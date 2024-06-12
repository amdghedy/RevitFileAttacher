using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Windows.Interop;

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class AddRevitFiles : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Show the folder selection window
                var folderSelectionWindow = new FolderSelectionWindow();
                var helper = new WindowInteropHelper(folderSelectionWindow);
                helper.Owner = commandData.Application.MainWindowHandle;

                if (folderSelectionWindow.ShowDialog() == true)
                {
                    string selectedPath = folderSelectionWindow.SelectedFolderPath;

                    // Get all Revit files in the selected directory
                    string[] revitFiles = Directory.GetFiles(selectedPath, "*.rvt", SearchOption.AllDirectories);

                    if (revitFiles.Length == 0)
                    {
                        TaskDialog.Show("No Files", "No Revit files found in the selected directory.");
                        return Result.Succeeded;
                    }

                    // Get the current document
                    UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                    Document doc = uiDoc.Document;

                    // Counter for successfully linked files
                    int filesLinked = 0;

                    // Iterate through each Revit file and link them
                    foreach (string file in revitFiles)
                    {
                        try
                        {
                            LinkRevitFile(doc, file);
                            filesLinked++;
                        }
                        catch (Exception ex)
                        {
                            // Log or handle specific file link error if necessary
                        }
                    }

                    TaskDialog.Show("Success", $"{filesLinked} Revit files linked successfully.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void LinkRevitFile(Document doc, string filePath)
        {
            using (Transaction trans = new Transaction(doc, "Link Revit File"))
            {
                trans.Start();

                // Create the ModelPath for the external file
                ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);

                // Link the Revit file
                RevitLinkOptions linkOptions = new RevitLinkOptions(false);
                ExternalResourceType resourceType = ExternalResourceTypes.BuiltInExternalResourceTypes.RevitLink;
                ExternalResourceReference extRef = ExternalResourceReference.CreateLocalResource(doc, resourceType, modelPath, PathType.Absolute);

                // Corrected to use the LinkLoadResult
                LinkLoadResult linkLoadResult = RevitLinkType.Create(doc, extRef, linkOptions);
                RevitLinkType linkType = doc.GetElement(linkLoadResult.ElementId) as RevitLinkType;

                // Create Revit link instance
                RevitLinkInstance.Create(doc, linkType.Id);

                trans.Commit();
            }
        }
    }
}
