using Microsoft.UI.Input;
using Microsoft.UI.Xaml.Controls;

namespace MailMind;

public sealed class ColumnSplitter : UserControl
{
    public ColumnSplitter()
    {
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
        Content = new Border { Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.Transparent) };
    }
}
