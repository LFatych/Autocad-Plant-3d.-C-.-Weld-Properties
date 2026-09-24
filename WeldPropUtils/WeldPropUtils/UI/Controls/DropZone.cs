using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WeldPropUtils.UI.Controls
{
    // A weld property field that accepts a dropped PropertyChip. On drop it runs DropCommand with the
    // dropped property name. Clicking it runs Command (assign the selected property).
    // IsMapped / IsDragOver drive the look (implicit style for DropZone in UI/Themes/SpdsStyles.xaml).
    public class DropZone : Button
    {
        public static readonly DependencyProperty TargetNameProperty =
            DependencyProperty.Register(nameof(TargetName), typeof(string), typeof(DropZone));

        public static readonly DependencyProperty SourceNameProperty =
            DependencyProperty.Register(nameof(SourceName), typeof(string), typeof(DropZone),
                new PropertyMetadata(null, (d, e) => ((DropZone)d).IsMapped = !string.IsNullOrEmpty(e.NewValue as string)));

        public static readonly DependencyProperty IsMappedProperty =
            DependencyProperty.Register(nameof(IsMapped), typeof(bool), typeof(DropZone));

        public static readonly DependencyProperty IsDragOverProperty =
            DependencyProperty.Register(nameof(IsDragOver), typeof(bool), typeof(DropZone));

        public static readonly DependencyProperty DropCommandProperty =
            DependencyProperty.Register(nameof(DropCommand), typeof(ICommand), typeof(DropZone));

        public DropZone()
        {
            AllowDrop = true;
        }

        // Weld property this field writes (e.g. "Material1").
        public string TargetName
        {
            get => (string)GetValue(TargetNameProperty);
            set => SetValue(TargetNameProperty, value);
        }

        // Part property mapped to it, or null.
        public string SourceName
        {
            get => (string)GetValue(SourceNameProperty);
            set => SetValue(SourceNameProperty, value);
        }

        public bool IsMapped
        {
            get => (bool)GetValue(IsMappedProperty);
            private set => SetValue(IsMappedProperty, value);
        }

        public bool IsDragOver
        {
            get => (bool)GetValue(IsDragOverProperty);
            private set => SetValue(IsDragOverProperty, value);
        }

        public ICommand DropCommand
        {
            get => (ICommand)GetValue(DropCommandProperty);
            set => SetValue(DropCommandProperty, value);
        }

        protected override void OnDragEnter(DragEventArgs e)
        {
            base.OnDragEnter(e);
            IsDragOver = e.Data.GetDataPresent(DataFormats.StringFormat);
        }

        protected override void OnDragOver(DragEventArgs e)
        {
            base.OnDragOver(e);
            e.Effects = e.Data.GetDataPresent(DataFormats.StringFormat) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        protected override void OnDragLeave(DragEventArgs e)
        {
            base.OnDragLeave(e);
            IsDragOver = false;
        }

        protected override void OnDrop(DragEventArgs e)
        {
            base.OnDrop(e);
            IsDragOver = false;
            if (e.Data.GetData(DataFormats.StringFormat) is string dropped && DropCommand?.CanExecute(dropped) == true)
                DropCommand.Execute(dropped);
        }
    }
}
