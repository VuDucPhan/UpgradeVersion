using Autodesk.Revit.Attributes ;
using Nice3point.Revit.Toolkit.External ;
using System.IO ;
using System.Reflection ;
using System.Windows ;
using Autodesk.Revit.UI ;
using MessageBox = System.Windows.MessageBox ;
using Microsoft.WindowsAPICodePack.Dialogs ;
using Autodesk.Revit.UI.Events ;

namespace UpgradeVersion.Commands ;

/// <summary>
///     External command entry point invoked from the Revit interface
/// </summary>
[UsedImplicitly]
[Transaction( TransactionMode.Manual )]
public class UpgradeVersionCommand : ExternalCommand
{
  public const string PlaceHolderName = "placeholder.rvt" ;

  public string PlaceHolderPath ;
  private FileLogger _logger ;

  public override void Execute()
  {
    var executingAssemblyLocation = Path.GetDirectoryName( Assembly.GetExecutingAssembly().Location ) ;
    PlaceHolderPath = Path.Combine( executingAssemblyLocation!, PlaceHolderName ) ;
    if ( ! File.Exists( PlaceHolderPath ) ) {
      var directoryName = Path.GetDirectoryName( PlaceHolderPath ) ;
      if ( directoryName != null && ! Directory.Exists( directoryName ) )
        Directory.CreateDirectory( directoryName ) ;

      var placeHolderDoc = ExternalCommandData.Application.Application.NewProjectDocument( UnitSystem.Metric ) ;
      placeHolderDoc.SaveAs( PlaceHolderPath ) ;
    }

    // Register DialogBoxShowing event
    ExternalCommandData.Application.DialogBoxShowing += ApplicationOnDialogBoxShowing ;

    try {
      // Show folder selection dialog
      using var folderDialog = new CommonOpenFileDialog { IsFolderPicker = true, Title = "Select folder containing Revit files to upgrade", InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) } ;

      if ( folderDialog.ShowDialog() != CommonFileDialogResult.Ok ) {
        return ;
      }

      var selectedFolder = folderDialog.FileName ;

      // Initialize logger
      _logger = new FileLogger( selectedFolder ) ;
      _logger.Log( "Starting upgrade process..." ) ;

      var uiApp = ExternalCommandData.Application ;
      var revitFiles = GetRevitFiles( selectedFolder ) ;

      UpdateProgress( $"Found {revitFiles.Count} files to upgrade" ) ;

      for ( int i = 0 ; i < revitFiles.Count ; i++ ) {
        var revitFile = revitFiles[ i ] ;
        try {
          UpdateProgress( $"Processing file {i + 1}/{revitFiles.Count}: {revitFile}" ) ;

          // Open and save document
          uiApp.OpenAndActivateDocument( revitFile ) ;
          CloseCurrentView( ExternalCommandData ) ;

          UpdateProgress( $"Completed file {i + 1}/{revitFiles.Count}: {revitFile}" ) ;
        }
        catch ( Exception ex ) {
          UpdateProgress( $"Error processing file {revitFile}: {ex.Message}" ) ;
          MessageBox.Show( $"Error processing file {revitFile}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error ) ;
        }
      }

      UpdateProgress( "Upgrade process completed!" ) ;
    }
    finally {
      // Unregister DialogBoxShowing event
      ExternalCommandData.Application.DialogBoxShowing -= ApplicationOnDialogBoxShowing ;
      _logger?.Dispose() ;
    }
  }

  private void UpdateProgress( string message )
  {
    _logger?.Log( message ) ;
  }

  private void ApplicationOnDialogBoxShowing( object sender, DialogBoxShowingEventArgs e )
  {
    if ( e.DialogId is "TaskDialog_Missing_Third_Party_Updater" or "TaskDialog_Missing_Third_Party_Updaters" )
      e.OverrideResult( 2 ) ;
    else if ( e.DialogId == "TaskDialog_Schema_Conflict" )
      e.OverrideResult( 1 ) ;
    else
      return ;

    if ( e.Cancellable )
      e.Cancel() ;
  }

  private List<string> GetRevitFiles( string folderPath )
  {
    // Get all .rvt files in the specified folder and all subfolders
    var allRevitFiles = Directory.GetFiles( folderPath, "*.rvt", SearchOption.AllDirectories ) ;

    // Filter out backup files (files ending with .0001, .0002, .0003, .0001.0001, etc.)
    var revitFiles = allRevitFiles.Where( file =>
    {
      var fileName = Path.GetFileNameWithoutExtension( file ) ;
      // Skip files that end with backup patterns (.0001, .0002, .0003, .0001.0001, etc.)
      return ! fileName.Contains( PlaceHolderName ) && ! System.Text.RegularExpressions.Regex.IsMatch( fileName, @"\.\d{4}(\.\d{4})?$" ) ;
    } ).ToList() ;

    return revitFiles ;
  }

  public void CloseCurrentView( ExternalCommandData commandData )
  {
    var uiDocument = commandData.Application.ActiveUIDocument ;
    var documentActiveView = uiDocument.ActiveView.Document ;
    commandData.Application.OpenAndActivateDocument( PlaceHolderPath ) ;
    documentActiveView.Close( true ) ;
  }
}