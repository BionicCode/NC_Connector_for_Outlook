// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.IO;

namespace NcTalkOutlookAddIn.Services
{
    // Commits a prepared file while preserving the last known-good primary as backup.
    internal static class DurableFileReplace
    {
        internal static bool CommitPreparedFile(
            string temporaryPath,
            string primaryPath,
            string backupPath)
        {
            if (!File.Exists(primaryPath))
            {
                File.Move(temporaryPath, primaryPath);
                File.Copy(primaryPath, backupPath, true);
                return true;
            }

            string preservedPath =
                backupPath
                + "."
                + Guid.NewGuid().ToString("N")
                + ".preserved";
            try
            {
                CopyAndVerify(primaryPath, preservedPath);
                File.Replace(
                    temporaryPath,
                    primaryPath,
                    backupPath,
                    true);
                return true;
            }
            catch (PlatformNotSupportedException)
            {
                CommitWithPreservedCopyFallback(
                    temporaryPath,
                    primaryPath,
                    backupPath,
                    preservedPath);
                return false;
            }
            catch (IOException)
            {
                CommitWithPreservedCopyFallback(
                    temporaryPath,
                    primaryPath,
                    backupPath,
                    preservedPath);
                return false;
            }
            finally
            {
                if (File.Exists(preservedPath))
                {
                    File.Delete(preservedPath);
                }
            }
        }

        private static void CommitWithPreservedCopyFallback(
            string temporaryPath,
            string primaryPath,
            string backupPath,
            string preservedPath)
        {
            // The order is intentional: the pre-replace primary is made durable
            // before the prepared replacement is allowed to touch the primary.
            CopyAndVerify(preservedPath, backupPath);
            if (File.Exists(temporaryPath))
            {
                File.Copy(temporaryPath, primaryPath, true);
            }
        }

        private static void CopyAndVerify(
            string sourcePath,
            string destinationPath)
        {
            File.Copy(sourcePath, destinationPath, true);
            var source = new FileInfo(sourcePath);
            var destination = new FileInfo(destinationPath);
            if (!destination.Exists
                || destination.Length != source.Length)
            {
                throw new IOException(
                    "The durable state file copy could not be verified.");
            }
        }
    }
}
