using System.IO;
using System.IO.Compression;
using RequirementTemplateGenerator;
using Xunit;

namespace RequirementTemplateGenerator.Tests
{
    public class DotxFixerTests
    {
        [Fact]
        public void GeneratedDotx_IsValidAfterFix()
        {
            var tmp = Path.Combine(Path.GetTempPath(), "req_template_test.dotx");
            if (File.Exists(tmp)) File.Delete(tmp);

            // Generate using the public API
            Program.GenerateTemplate(tmp);

            // Ensure file exists
            Assert.True(File.Exists(tmp));

            // Inspect ZIP contents
            using (var z = new ZipArchive(File.OpenRead(tmp)))
            {
                // Check expected parts exist
                Assert.NotNull(z.GetEntry("word/document.xml"));
                Assert.NotNull(z.GetEntry("word/styles.xml"));
                Assert.NotNull(z.GetEntry("word/numbering.xml"));
                Assert.NotNull(z.GetEntry("[Content_Types].xml"));
                Assert.NotNull(z.GetEntry("_rels/.rels"));
                Assert.NotNull(z.GetEntry("word/_rels/document.xml.rels"));

                // Check Content_Types contains Override for /word/document.xml
                var ct = z.GetEntry("[Content_Types].xml");
                using (var r = new StreamReader(ct.Open()))
                {
                    var text = r.ReadToEnd();
                    Assert.Contains("/word/document.xml", text);
                    Assert.Contains("/word/styles.xml", text);
                    Assert.Contains("/word/numbering.xml", text);
                }

                // Check rel targets do NOT start with a leading slash
                var rel = z.GetEntry("_rels/.rels");
                using (var r = new StreamReader(rel.Open()))
                {
                    var text = r.ReadToEnd();
                    Assert.DoesNotContain("Target=\"/word/", text);
                    Assert.Contains("Target=\"word/document.xml\"", text);
                }
            }

            // Clean up
            File.Delete(tmp);
        }
    }
}
