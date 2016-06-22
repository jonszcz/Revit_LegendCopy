using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Media.Imaging;

namespace Elk
{
    [Transaction(TransactionMode.Manual)]
    public class LegendCopyCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;

                // Get the current selection and make sure there's only one item selected.
                Selection sel = uiDoc.Selection;
                if (sel.GetElementIds().Count == 1)
                {
                    // Get the selected element
                    Element e = uiDoc.Document.GetElement(sel.GetElementIds().First());

                    // Make sure the element was retrieved and that it's a viewport element
                    if (e != null && e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Viewports)
                    {
                        // Get the viewport
                        Viewport vp = e as Viewport;

                        // Get the legend view's name
                        string viewName = vp.get_Parameter(BuiltInParameter.VIEWPORT_VIEW_NAME).AsString();

                        // Get the viewport type
                        ElementId vpTypeId = vp.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsElementId();
                        View view = uiDoc.Document.GetElement(vp.ViewId) as View;

                        // Make sure the view is a legend view
                        if (view.ViewType == ViewType.Legend)
                        {
                            // Get the location of the viewport
                            XYZ loc = vp.GetBoxCenter();

                            FilteredElementCollector sheetCollector = new FilteredElementCollector(uiDoc.Document);
                            sheetCollector.OfClass(typeof(ViewSheet));
                            List<ViewSheet> sheets = new List<ViewSheet>();
                            foreach (ViewSheet vs in sheetCollector)
                            {
                                sheets.Add(vs);
                            }

                            // Pass the location, view, and viewport type to the form
                            System.Diagnostics.Process proc = System.Diagnostics.Process.GetCurrentProcess();
                            IntPtr handle = proc.MainWindowHandle;

                            LegendCopyForm form = new LegendCopyForm(uiDoc.Document, view, vpTypeId, loc, sheets);
                            System.Windows.Interop.WindowInteropHelper wih = new System.Windows.Interop.WindowInteropHelper(form);
                            wih.Owner = handle;
                            form.ShowDialog();
                        }
                        else
                            TaskDialog.Show("Warning - 2", "Make sure only one Legend Viewport is selected before running this command.");
                    }
                    else
                        TaskDialog.Show("Warning - 1", "Make sure only one Legend Viewport is selected before running this command.");
                }
                else
                    TaskDialog.Show("Warning - 0", "Make sure only one Legend Viewport is selected before running this command.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public class LegendCopyApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string path = typeof(LegendCopyApp).Assembly.Location;

            // Create the PushButtonData
            PushButtonData legendCopyPBD = new PushButtonData(
                "Legend Copy", "Legend\nCopy", path, "Elk.LegendCopyCmd")
            {
                LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Legend_32x32.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()),
                ToolTip = "Copy a legend view to multiple sheets in the same location.",
            };

            // Add the button to the ribbon
            RevitCommon.UI.AddToRibbon(application, Properties.Settings.Default.TabName, Properties.Settings.Default.PanelName, legendCopyPBD);

            return Result.Succeeded;
        }
    }
}
