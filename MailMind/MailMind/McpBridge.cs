using System;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;

namespace MailMind;

// The stdio process forwards requests to the running UI, which owns the mail and draft state.
internal static class McpBridge
{
    private const string PipeName = "MailMind.Mcp.v1";
    private static CancellationTokenSource? hostCancellation;
    private static readonly JsonSerializerOptions Json = MailAgent.ReadableJson;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetNamedPipeServerProcessId(Microsoft.Win32.SafeHandles.SafePipeHandle pipe, out uint processId);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AllowSetForegroundWindow(uint processId);

    internal static void StartHost(DispatcherQueue dispatcher, Func<string, JsonElement, CancellationToken, Task<string>> execute)
    {
        hostCancellation = new CancellationTokenSource();
        _ = Task.Run(() => HostAsync(dispatcher, execute, hostCancellation.Token));
    }

    internal static void StopHost() => hostCancellation?.Cancel();

    private static async Task HostAsync(DispatcherQueue dispatcher, Func<string, JsonElement, CancellationToken, Task<string>> execute, CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            var pipe = new NamedPipeServerStream(PipeName, PipeDirection.InOut, NamedPipeServerStream.MaxAllowedServerInstances, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
            try { await pipe.WaitForConnectionAsync(token); }
            catch (OperationCanceledException) { pipe.Dispose(); break; }
            catch { pipe.Dispose(); break; }
            _ = HandleConnectionAsync(pipe, dispatcher, execute, token);
        }
    }

    private static async Task HandleConnectionAsync(NamedPipeServerStream pipe, DispatcherQueue dispatcher, Func<string, JsonElement, CancellationToken, Task<string>> execute, CancellationToken token)
    {
        using (pipe)
        using (var reader = new StreamReader(pipe, Encoding.UTF8, false, 4096, true))
        using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
        {
            try
            {
                string? line;
                while ((line = await reader.ReadLineAsync(token)) != null)
                {
                    string response = await DispatchAsync(line, async (name, arguments) =>
                    {
                        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                        if (!dispatcher.TryEnqueue(async () =>
                        {
                            try { completion.TrySetResult(await execute(name, arguments, token)); }
                            catch (Exception ex) { completion.TrySetException(ex); }
                        })) throw new InvalidOperationException("MailMind is closing.");
                        return await completion.Task;
                    });
                    if (response.Length != 0) await writer.WriteLineAsync(response.AsMemory(), token);
                }
            }
            catch (OperationCanceledException) { }
            catch (IOException) { }
        }
    }

    internal static async Task RunStdioAsync()
    {
        using var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false), false);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        using var pipe = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        try { await pipe.ConnectAsync(3000); }
        catch (Exception ex) { await Console.Error.WriteLineAsync("Start MailMind before connecting to its MCP server: " + ex.Message); return; }
        using var reader = new StreamReader(pipe, Encoding.UTF8, false, 4096, true);
        using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true };
        string? line;
        while ((line = await stdin.ReadLineAsync()) != null)
        {
            // Pass foreground permission to the window before it handles a draft request.
            try
            {
                using var request = JsonDocument.Parse(line);
                var root = request.RootElement;
                if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("method", out var method) && method.ValueKind == JsonValueKind.String && method.GetString() == "tools/call" &&
                    root.TryGetProperty("params", out var parameters) && parameters.ValueKind == JsonValueKind.Object && parameters.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String &&
                    name.GetString() is "compose_email" or "update_draft" or "add_attachment" && GetNamedPipeServerProcessId(pipe.SafePipeHandle, out uint processId))
                    AllowSetForegroundWindow(processId);
            }
            catch (JsonException) { }
            await writer.WriteLineAsync(line);
            // Notifications have no response.
            bool hasId;
            try { using var request = JsonDocument.Parse(line); var root = request.RootElement; hasId = root.ValueKind != JsonValueKind.Object || root.TryGetProperty("id", out _) || !root.TryGetProperty("jsonrpc", out var version) || version.ValueKind != JsonValueKind.String || version.GetString() != "2.0" || !root.TryGetProperty("method", out var method) || method.ValueKind != JsonValueKind.String; }
            catch (JsonException) { hasId = true; }
            if (!hasId) continue;
            string? response = await reader.ReadLineAsync();
            if (response == null) break;
            await stdout.WriteLineAsync(response);
        }
    }

    internal static async Task<string> CallToolAsync(string name, JsonElement arguments, CancellationToken token)
    {
        using var pipe = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        await pipe.ConnectAsync(3000, token);
        using var reader = new StreamReader(pipe, Encoding.UTF8, false, 4096, true);
        using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true };
        await writer.WriteLineAsync(JsonSerializer.Serialize(new { jsonrpc = "2.0", id = 1, method = "tools/call", @params = new { name, arguments } }, Json).AsMemory(), token);
        string line = await reader.ReadLineAsync(token) ?? throw new IOException("MCP connection closed.");
        using var response = JsonDocument.Parse(line);
        if (response.RootElement.TryGetProperty("error", out var error)) throw new InvalidOperationException(error.GetProperty("message").GetString());
        return response.RootElement.GetProperty("result").GetProperty("content")[0].GetProperty("text").GetString() ?? "";
    }

    private static async Task<string> DispatchAsync(string line, Func<string, JsonElement, Task<string>> execute)
    {
        JsonElement id = default;
        bool hasId = false;
        string Error(int code, string message) => JsonSerializer.Serialize(new { jsonrpc = "2.0", id = hasId ? (object)id : null, error = new { code, message } }, Json);
        try
        {
            using var document = JsonDocument.Parse(line);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return Error(-32600, "Invalid request");
            hasId = root.TryGetProperty("id", out id);
            if (!root.TryGetProperty("jsonrpc", out var version) || version.GetString() != "2.0" || !root.TryGetProperty("method", out var methodValue) || methodValue.ValueKind != JsonValueKind.String)
                return Error(-32600, "Invalid request");
            string method = methodValue.GetString()!;
            if (!hasId) return "";
            object result;
            switch (method)
            {
                case "initialize":
                    string protocol = "2024-11-05";
                    if (root.TryGetProperty("params", out var init) && init.ValueKind == JsonValueKind.Object && init.TryGetProperty("protocolVersion", out var requested) && requested.ValueKind == JsonValueKind.String && (requested.GetString() == "2025-03-26" || requested.GetString() == "2025-06-18")) protocol = requested.GetString()!;
                    result = new { protocolVersion = protocol, capabilities = new { tools = new { } }, serverInfo = new { name = "MailMind", version = "1.0.0" } };
                    break;
                case "ping": result = new { }; break;
                case "tools/list":
                    result = new { tools = MailAgent.ToolDefinitions().Select(tool => { var item = JsonSerializer.SerializeToElement(tool, Json).GetProperty("function"); return new { name = item.GetProperty("name").GetString(), description = item.GetProperty("description").GetString(), inputSchema = item.GetProperty("parameters").Clone() }; }).ToArray() };
                    break;
                case "tools/call":
                    if (!root.TryGetProperty("params", out var parameters) || parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty("name", out var nameValue) || nameValue.ValueKind != JsonValueKind.String) return Error(-32602, "Invalid tool parameters");
                    string name = nameValue.GetString()!;
                    if (!MailAgent.Skills.ContainsKey(name)) return Error(-32602, "Unknown tool");
                    JsonElement arguments = parameters.TryGetProperty("arguments", out var arg) ? arg.Clone() : JsonSerializer.SerializeToElement(new { });
                    if (arguments.ValueKind != JsonValueKind.Object) return Error(-32602, "Arguments must be an object");
                    string output; bool failed = false;
                    try { output = await execute(name, arguments); using var parsed = JsonDocument.Parse(output); failed = parsed.RootElement.ValueKind == JsonValueKind.Object && parsed.RootElement.TryGetProperty("success", out var success) && success.ValueKind == JsonValueKind.False; }
                    catch (Exception ex) { output = JsonSerializer.Serialize(new { success = false, error = ex.Message }, Json); failed = true; }
                    result = new { content = new[] { new { type = "text", text = output } }, isError = failed };
                    break;
                default: return Error(-32601, "Method not found");
            }
            return JsonSerializer.Serialize(new { jsonrpc = "2.0", id, result }, Json);
        }
        catch (JsonException) { return Error(-32700, "Parse error"); }
        catch (Exception ex) { return Error(-32600, ex.Message); }
    }
}
