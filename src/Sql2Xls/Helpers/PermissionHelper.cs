using System.Runtime.InteropServices;
using System.Security.AccessControl;

namespace Sql2Xls.Helpers;

public static class PermissionHelper
{
    public static bool CheckWriteAccessToFolder(string folderPath)
    {

        var di = new DirectoryInfo(folderPath);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            try
            {
                // Attempt to get a list of security permissions from the folder. 
                // This will raise an exception if the path is read only or do not have access to view the permissions. 

                DirectorySecurity ds = di.GetAccessControl();
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }


        var mode = di.UnixFileMode;
        if (mode.HasFlag(UnixFileMode.UserRead) && mode.HasFlag(UnixFileMode.UserWrite))
        {
            return true;
        }

        return false;

        /*
        try
        {
            File.SetUnixFileMode(folderPath, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        */
    }
}
