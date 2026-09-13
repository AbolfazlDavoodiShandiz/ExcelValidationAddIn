using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using ExcelValidationAddIn.Models;
using Newtonsoft.Json;

namespace ExcelValidationAddIn.Services
{
    /// <summary>
    /// Stores the session as JSON, encrypted with DPAPI (CurrentUser scope) so
    /// the JWT is not sitting in a plain-text file under %AppData%.
    /// </summary>
    public static class SettingsStore
    {
        private static readonly string FolderPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelValidationAddIn");

        private static readonly string FilePath = Path.Combine(FolderPath, "session.dat");

        public static Session Load()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    return null;
                }

                byte[] encrypted = File.ReadAllBytes(FilePath);
                byte[] plainBytes = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                string json = Encoding.UTF8.GetString(plainBytes);
                return JsonConvert.DeserializeObject<Session>(json);
            }
            catch
            {
                // Corrupted or unreadable (e.g. written by a different Windows
                // user) - treat as "no session" rather than throwing.
                return null;
            }
        }

        public static void Save(Session session)
        {
            Directory.CreateDirectory(FolderPath);
            string json = JsonConvert.SerializeObject(session);
            byte[] plainBytes = Encoding.UTF8.GetBytes(json);
            byte[] encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(FilePath, encrypted);
        }

        public static void Clear()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }
            }
            catch
            {
                // Best-effort - nothing user-facing to do if this fails.
            }
        }
    }
}
