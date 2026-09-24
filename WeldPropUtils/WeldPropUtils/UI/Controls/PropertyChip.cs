using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WeldPropUtils.UI.Controls
{
    // A draggable property of the connected parts. Click selects it (Command); dragging starts a
    // drag-and-drop with PropertyName as text, which a DropZone accepts.
    // Look: implicit style for PropertyChip in UI/Themes/SpdsStyles.xaml.
    public class PropertyChip : Button
    {
        public static readonly DependencyProperty PropertyNameProperty =
            DependencyProperty.Register(nameof(PropertyName), typeof(string), typeof(PropertyChip));

        public static readonly DependencyProperty ScopeProperty =
            DependencyProperty.Register(nameof(Scope), typeof(string), typeof(PropertyChip));

        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(nameof(IsSelected), typeof(bool), typeof(PropertyChip));

        private Point _dragStart;
        private bool _pressed;

        // Internal name of the part property (e.g. "MatchingPipeOd").
        public string PropertyName
        {
            get => (string)GetValue(PropertyNameProperty);
            set => SetValue(PropertyNameProperty, value);
        }

        // Which class groups have it (e.g. "all classes", "Pipe, Fittings").
        public string Scope
        {
            get => (string)GetValue(ScopeProperty);
            set => SetValue(ScopeProperty, value);
        }

        public bool IsSelected
        {
            get => (bool)GetValue(IsSelectedProperty);
            set => SetValue(IsSelectedProperty, value);
        }

        protected override void OnPreviewMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnPreviewMouseLeftButtonDown(e);
            _dragStart = e.GetPosition(this);
            _pressed = true;
        }

        protected override void OnPreviewMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            _pressed = false;
            base.OnPreviewMouseLeftButtonUp(e);
        }

        protected override void OnPreviewMouseMove(MouseEventArgs e)
        {
            base.OnPreviewMouseMove(e);
            if (!_pressed || e.LeftButton != MouseButtonState.Pressed || string.IsNullOrEmpty(PropertyName)) return;
            Vector moved = e.GetPosition(this) - _dragStart;
            if (Math.Abs(moved.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(moved.Y) < SystemParameters.MinimumVerticalDragDistance) return;
            _pressed = false;
            DragDrop.DoDragDrop(this, new DataObject(DataFormats.StringFormat, PropertyName), DragDropEffects.Copy);
        }
    }
}
