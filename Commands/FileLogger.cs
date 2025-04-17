using System.IO ;

namespace UpgradeVersion.Commands ;

public class FileLogger
{
  private readonly string _logFilePath ;
  private readonly StreamWriter _writer ;

  public FileLogger( string folderPath )
  {
    _logFilePath = Path.Combine( folderPath, "upgrade_log.txt" ) ;
    _writer = new StreamWriter( _logFilePath, true ) { AutoFlush = true } ;
  }

  public void Log( string message )
  {
    var timestamp = DateTime.Now.ToString( "yyyy-MM-dd HH:mm:ss" ) ;
    _writer.WriteLine( $"[{timestamp}] {message}" ) ;
  }

  public void Dispose()
  {
    _writer?.Dispose() ;
  }
}