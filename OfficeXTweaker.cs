using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace BypassAv
{
    public class OfficeXTweaker
    {
        private readonly string mInputFile;

        public OfficeXTweaker(string inputFile)
        {
            mInputFile = inputFile;
        }

        public void RemoveProperties(string outputPath)
        {
            if (outputPath != mInputFile)
            {
                File.Copy(mInputFile, outputPath, true);
            }

            // Open OfficeX as archive
            using ZipArchive archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            RemovePropertiesFiles(archive);

            // Get the main relation entry - "_rels\.rels"
            ZipArchiveEntry relsEntry = archive.GetEntry(Consts.TOP_RELATIONS);
            using Stream relsStream = relsEntry.Open();

            RemoveSettingsFromRelations(relsStream);
        }

        private static void RemoveSettingsFromRelations(Stream relsStream)
        {
            XDocument doc = XDocument.Load(relsStream);
            IEnumerable<XElement> relsToRemove = doc.Descendants()
                .Where(d => d.Name.LocalName == "Relationship" &&
                            Consts.PROPERTY_ENTRIES.Contains(d.Attribute("Target").Value));

            foreach (XElement remEl in relsToRemove.ToArray())
            {
                remEl.Remove();
            }

            relsStream.Position = 0;
            doc.Save(relsStream);
            relsStream.SetLength(relsStream.Position);
        }

        private void RemovePropertiesFiles(ZipArchive archive)
        {
            foreach (ZipArchiveEntry entryToRemove in Consts.PROPERTY_ENTRIES.Select(ent => archive.GetEntry(ent)).Where(e => e != null))
            {
                entryToRemove.Delete();
            }
        }
    }
}
