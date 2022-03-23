using System.IO;
using System.Management;
using OfficeOpenXml;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public static class FileManagerController
    {
        private static void Delete(string file)
        {
            if (File.Exists(file))
                File.Delete(file);
        }

        private static void CreateDirectory(string directory)
        {
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }
        
        private static void FileExists(string file)
        {
            Delete(file);
        }

        private static void DirectoryExists(string directory)
        {
            CreateDirectory(directory);
        }
        
        public static string GetFile(string attachments, string name, string format)
        {
            var file = @$"{GetDirectory(attachments)}{name}.{format}";
            FileExists(file);
            return file;
        }

        private static string GetDirectory(string attachments)
        {
            var directory = @$"{attachments}\";
            DirectoryExists(directory);
            return directory;
        }
    }
}