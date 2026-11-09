using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;
using MimeKit;
using Windows.System;

namespace MailMind;

public sealed class RecipientBox : UserControl
{
    private readonly StackPanel pills = new() { Orientation = Orientation.Horizontal, Spacing = 4 };
    private readonly AutoSuggestBox input = new() { MinWidth = 130, Width = 180, BorderThickness = new Thickness(0), Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent) };
    private readonly List<MailboxAddress> recipients = new();
    private bool updating;
    public Func<string, List<ContactsPane.Contact>>? Search { get; set; }

    public RecipientBox()
    {
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        row.Children.Add(pills); row.Children.Add(input);
        var border = (Border)XamlReader.Load("<Border xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' CornerRadius='5' BorderThickness='1' BorderBrush='{ThemeResource DividerBrush}' Background='{ThemeResource SurfaceBrush}' />");
        border.Child = new ScrollViewer { Content = row, HorizontalScrollMode = ScrollMode.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden, VerticalScrollBarVisibility = ScrollBarVisibility.Disabled };
        Content = border;
        input.ItemTemplate = (DataTemplate)XamlReader.Load("<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'><StackPanel Orientation='Horizontal' Spacing='12'><TextBlock Text='{Binding Name}' FontWeight='SemiBold'/><TextBlock Text='{Binding Email}' Opacity='0.7'/></StackPanel></DataTemplate>");
        input.TextChanged += InputChanged;
        input.SuggestionChosen += (_, e) => { if (e.SelectedItem is ContactsPane.Contact contact) { Add(new MailboxAddress(contact.Name, contact.Email)); ClearInput(); } };
        input.QuerySubmitted += (_, e) => { if (e.ChosenSuggestion == null) CommitInput(); };
        input.KeyDown += (_, e) => { if (e.Key == VirtualKey.Back && input.Text.Length == 0 && recipients.Count > 0) { recipients.RemoveAt(recipients.Count - 1); Render(); e.Handled = true; } };
    }
    public string Text
    {
        get => string.Join(", ", recipients.Select(r => r.ToString()).Concat(string.IsNullOrWhiteSpace(input.Text) ? Array.Empty<string>() : new[] { input.Text.Trim() }));
        set
        {
            recipients.Clear(); ClearInput();
            if (!string.IsNullOrWhiteSpace(value))
            {
                if (InternetAddressList.TryParse(value.Replace(';', ','), out var addresses))
                    foreach (var address in addresses.Mailboxes) Add(Resolve(address));
                else input.Text = value;
            }
            Render();
        }
    }
    private MailboxAddress Resolve(MailboxAddress address)
    {
        var contact = Search?.Invoke(address.Address).FirstOrDefault(c => string.Equals(c.Email, address.Address, StringComparison.OrdinalIgnoreCase));
        return contact == null ? address : new MailboxAddress(contact.Name, contact.Email);
    }
    private void InputChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs e)
    {
        if (updating || e.Reason != AutoSuggestionBoxTextChangeReason.UserInput) return;
        int separator;
        while ((separator = sender.Text.IndexOfAny(new[] { ',', ';', '，', '；' })) >= 0)
        {
            string part = sender.Text.Substring(0, separator).Trim();
            if (MailboxAddress.TryParse(part, out var address) && address.Address.Contains('@'))
            {
                Add(Resolve(address));
                updating = true; sender.Text = sender.Text.Substring(separator + 1).TrimStart(); updating = false;
            }
            else break;
        }
        sender.ItemsSource = Search?.Invoke(sender.Text) ?? new List<ContactsPane.Contact>();
        sender.IsSuggestionListOpen = sender.ItemsSource is System.Collections.ICollection items && items.Count > 0;
    }
    private void CommitInput()
    {
        if (MailboxAddress.TryParse(input.Text.Trim(), out var address) && address.Address.Contains('@')) { Add(Resolve(address)); ClearInput(); }
    }
    private void ClearInput() { updating = true; input.Text = ""; input.IsSuggestionListOpen = false; updating = false; }
    private void Add(MailboxAddress address)
    {
        if (recipients.Any(r => string.Equals(r.Address, address.Address, StringComparison.OrdinalIgnoreCase))) return;
        recipients.Add(address); Render();
    }
    private void Render()
    {
        pills.Children.Clear();
        foreach (var recipient in recipients.ToArray())
        {
            string label = string.IsNullOrWhiteSpace(recipient.Name) ? recipient.Address.Split('@')[0] : recipient.Name;
            var button = new Button { Padding = new Thickness(10, 3, 8, 3), CornerRadius = new CornerRadius(16), Content = label + "  ×", Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 225, 235, 252)), Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 25, 67, 140)) };
            ToolTipService.SetToolTip(button, recipient.Address);
            button.Click += (_, _) => { recipients.Remove(recipient); Render(); input.Focus(FocusState.Programmatic); };
            pills.Children.Add(button);
        }
    }
}
