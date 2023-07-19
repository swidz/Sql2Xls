using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sql2Xls.Excel;

public class ExcelDocument : IDisposable
{
    private SpreadsheetDocument xlDocument;
    
    private bool disposed = false;

    public virtual void Close()
    {
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposed)
            return;

        if (disposing)
        {
            // Free any other managed objects here.
            Close();

            if (xlDocument != null)
            {
                xlDocument.Dispose();
                xlDocument = null;
            }
        }

        // Free any unmanaged objects here.
        //
        disposed = true;
    }

    ~ExcelDocument()
    {
        Dispose(false);
    }
}
