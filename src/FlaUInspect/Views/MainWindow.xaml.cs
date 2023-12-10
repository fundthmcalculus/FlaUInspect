using System.Configuration;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using FlaUI.Core;
using FlaUInspect.ViewModels;

namespace FlaUInspect.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly MainViewModel _vm;

        public MainWindow()
        {
            InitializeComponent();
            AppendVersionToTitle();
            Height = 550;
            Width = 700;
            Loaded += MainWindow_Loaded;
            _vm = new MainViewModel();
            DataContext = _vm;
        }

        private void AppendVersionToTitle()
        {
            var attr = Assembly.GetEntryAssembly().GetCustomAttribute(typeof(AssemblyInformationalVersionAttribute)) as AssemblyInformationalVersionAttribute;
            if (attr != null)
            {
                Title += " v" + attr.InformationalVersion;
            }
        }

        private void MainWindow_Loaded(object sender, System.EventArgs e)
        {
            if (!_vm.IsInitialized)
            {
                // Start application if version is saved
                var version = ConfigurationManager.AppSettings["version"];
                if (version == "2")
                {
                    _vm.Initialize(AutomationType.UIA2);
                }
                else if (version == "3")
                {
                    _vm.Initialize(AutomationType.UIA3);
                }
                else
                {
                    var dlg = new ChooseVersionWindow { Owner = this };
                    if (dlg.ShowDialog() != true)
                    {
                        Close();
                    }
                    // Save selected UIA version if dialog is not closed
                    else if (dlg.SelectedAutomationType == AutomationType.UIA2 
                                || dlg.SelectedAutomationType == AutomationType.UIA3)
                    {
                        Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                        config.AppSettings.Settings.Remove("version");
                        config.AppSettings.Settings.Add("version", dlg.SelectedAutomationType == AutomationType.UIA2 ? "2" : "3");
                        config.Save(ConfigurationSaveMode.Minimal);
                    }

                    _vm.Initialize(dlg.SelectedAutomationType);
                }
                Loaded -= MainWindow_Loaded;
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TreeViewSelectedHandler(object sender, RoutedEventArgs e)
        {
            var item = sender as TreeViewItem;
            if (item != null)
            {
                item.BringIntoView();
                e.Handled = true;
            }
        }
        private void InvokePatternActionHandler(object sender, RoutedEventArgs e)
        {
            DetailViewModel vm = (DetailViewModel)((Button)sender).DataContext;
            if (vm.ActionToExecute != null)
            {
                Task.Run(() =>
                {
                    vm.ActionToExecute();
                });
            }
        }

        private void StackPanel_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // TODO - We should make this language flexible, as well as other configurable text
            // TODO - Have option to copy relevant extension method
            // _window.FindFirstDescendant(Function(x) x.ByNameOrAutomationId("")).AsButton()
            var boundItem = ((sender as StackPanel)?.DataContext as ElementViewModel);
            string automationNameQuery = "";
            if (boundItem != null)
            {
                automationNameQuery = boundItem.AutomationId ?? boundItem.Name ?? "";
                var codeString = $"_window.FindFirstDescendant(Function(x) x.ByNameOrAutomationId(\"{automationNameQuery}\"))";
                // If known type, parse that and return
                string asTypeCall = GetTypeCall(boundItem);
                if (!string.IsNullOrWhiteSpace(asTypeCall))
                    codeString = $"{codeString}.{asTypeCall}()";

                Clipboard.SetText(codeString);
            }
        }

        private static string GetTypeCall(ElementViewModel boundItem)
        {
            var asTypeCall = "";
            switch (boundItem.ControlType)
            {
                case FlaUI.Core.Definitions.ControlType.Button:
                    asTypeCall = "AsButton";
                    break;
                case FlaUI.Core.Definitions.ControlType.Unknown:
                    break;
                case FlaUI.Core.Definitions.ControlType.AppBar:
                    break;
                case FlaUI.Core.Definitions.ControlType.Calendar:
                    asTypeCall = "AsCalendar";
                    break;
                case FlaUI.Core.Definitions.ControlType.CheckBox:
                    asTypeCall = "AsCheckBox";
                    break;
                case FlaUI.Core.Definitions.ControlType.ComboBox:
                    asTypeCall = "AsComboBox";
                    break;
                case FlaUI.Core.Definitions.ControlType.Custom:
                    break;
                case FlaUI.Core.Definitions.ControlType.DataGrid:
                    asTypeCall = "AsDataGridView";
                    break;
                case FlaUI.Core.Definitions.ControlType.DataItem:
                    break;
                case FlaUI.Core.Definitions.ControlType.Document:
                    break;
                case FlaUI.Core.Definitions.ControlType.Edit:
                    asTypeCall = "AsTextBox";
                    break;
                case FlaUI.Core.Definitions.ControlType.Group:
                    break;
                case FlaUI.Core.Definitions.ControlType.Header:
                    asTypeCall = "AsGridHeader";
                    break;
                case FlaUI.Core.Definitions.ControlType.HeaderItem:
                    asTypeCall = "AsGridHeaderItem";
                    break;
                case FlaUI.Core.Definitions.ControlType.Hyperlink:
                    break;
                case FlaUI.Core.Definitions.ControlType.Image:
                    break;
                case FlaUI.Core.Definitions.ControlType.List:
                    asTypeCall = "AsListBox";
                    break;
                case FlaUI.Core.Definitions.ControlType.ListItem:
                    asTypeCall = "AsListBoxItem";
                    break;
                case FlaUI.Core.Definitions.ControlType.MenuBar:
                    break;
                case FlaUI.Core.Definitions.ControlType.Menu:
                    asTypeCall = "AsMenu";
                    break;
                case FlaUI.Core.Definitions.ControlType.MenuItem:
                    asTypeCall = "AsMenuItem";
                    break;
                case FlaUI.Core.Definitions.ControlType.Pane:
                    break;
                case FlaUI.Core.Definitions.ControlType.ProgressBar:
                    asTypeCall = "AsProgressBar";
                    break;
                case FlaUI.Core.Definitions.ControlType.RadioButton:
                    asTypeCall = "AsRadioButton";
                    break;
                case FlaUI.Core.Definitions.ControlType.ScrollBar:
                    break;
                case FlaUI.Core.Definitions.ControlType.SemanticZoom:
                    break;
                case FlaUI.Core.Definitions.ControlType.Separator:
                    break;
                case FlaUI.Core.Definitions.ControlType.Slider:
                    asTypeCall = "AsSlider";
                    break;
                case FlaUI.Core.Definitions.ControlType.Spinner:
                    asTypeCall = "AsSpinner";
                    break;
                case FlaUI.Core.Definitions.ControlType.SplitButton:
                    break;
                case FlaUI.Core.Definitions.ControlType.StatusBar:
                    break;
                case FlaUI.Core.Definitions.ControlType.Tab:
                    asTypeCall = "AsTab";
                    break;
                case FlaUI.Core.Definitions.ControlType.TabItem:
                    asTypeCall = "AsTabItem";
                    break;
                case FlaUI.Core.Definitions.ControlType.Table:
                    asTypeCall = "AsGrid";
                    break;
                case FlaUI.Core.Definitions.ControlType.Text:
                    asTypeCall = "AsLabel";
                    break;
                case FlaUI.Core.Definitions.ControlType.Thumb:
                    asTypeCall = "AsThumb";
                    break;
                case FlaUI.Core.Definitions.ControlType.TitleBar:
                    asTypeCall = "AsTitleBar";
                    break;
                case FlaUI.Core.Definitions.ControlType.ToolBar:
                    break;
                case FlaUI.Core.Definitions.ControlType.ToolTip:
                    break;
                case FlaUI.Core.Definitions.ControlType.Tree:
                    asTypeCall = "AsTree";
                    break;
                case FlaUI.Core.Definitions.ControlType.TreeItem:
                    asTypeCall = "AsTreeItem";
                    break;
                case FlaUI.Core.Definitions.ControlType.Window:
                    asTypeCall = "AsWindow";
                    break;
            }

            return asTypeCall;
        }
    }
}
