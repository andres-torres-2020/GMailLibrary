using Google.Apis.Gmail.v1.Data;
using System;
using System.IO;
using System.Text;
using ApplicationLogger;

namespace GMailLibrary
{
    public class EmailAttachmentProcessor : IGMailAttachmentProcessor
    {
        readonly string LocalDataPath;
        readonly ILogger AppLogger;
        public EmailAttachmentProcessor(string localDataPath, ILogger appLogger)
        {
            LocalDataPath = localDataPath;
            AppLogger = appLogger;
        }
        public Boolean Process(string attachmentFilename, MessagePartBody messagePartBody)
        {
            if (messagePartBody != null)
            {
                var decodedContents = FromBase64ForUrlString(messagePartBody.Data);
                return SaveBinaryFile(attachmentFilename, decodedContents);
            }
            return false;
        }
        public Boolean SaveBinaryFile(string fileName, byte[] fileContents)
        {
            try
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
                string extension = Path.GetExtension(fileName);
                string destinationFilename = LocalDataPath + fileName;

                // construct a file name for a new file
                int fileCounter = 1;
                while (File.Exists(destinationFilename))
                {
                    destinationFilename = LocalDataPath + fileNameWithoutExtension + "-" + fileCounter + extension;
                    fileCounter++;
                }

                // write the binary contents to a file
                AppLogger.Log($"SaveBinaryFile: saving [{fileName}] to [{destinationFilename}]");
                FileStream output = File.OpenWrite(destinationFilename);
                output.Write(fileContents, 0, fileContents.Length);
                output.Flush();
                output.Close();
                output.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Log($"SaveBinaryFile - error saving to binary file ({fileName})\n\t{ex.Message}");
                return false;
            }
        }
        public static byte[] FromBase64ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            StringBuilder result = new StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            result.Append(String.Empty.PadRight(padChars, '='));
            result.Replace('-', '+');
            result.Replace('_', '/');
            return Convert.FromBase64String(result.ToString());
        }
    }
}
