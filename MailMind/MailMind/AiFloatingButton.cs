using System.Numerics;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

namespace MailMind;

public sealed class AiFloatingButton : Button
{
    public AiFloatingButton()
    {
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.Hand);
        CenterPoint = new Vector3(30, 30, 0);
    }

    protected override void OnPointerEntered(PointerRoutedEventArgs e)
    {
        base.OnPointerEntered(e);
        Scale = new Vector3(1.08f, 1.08f, 1);
    }

    protected override void OnPointerExited(PointerRoutedEventArgs e)
    {
        base.OnPointerExited(e);
        Scale = Vector3.One;
    }
}
