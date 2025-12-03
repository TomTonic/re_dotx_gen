using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace RequirementTemplateGenerator
{
    public static class DotxFixer
    {
        /// <summary>
        /// Ensure .dotx has correct [Content_Types].xml and relationship targets.
        /// Rewrites the archive in-place (atomic replace).
        /// </summary>
        public static void FixDotx(string filePath)
        {
            // Determine if this is a .docx or .dotx to use correct content type
            bool isDocument = Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase);

            // First, inspect the archive to determine whether a fix is necessary
            bool needsFix = false;
            bool relsHaveLeadingSlash = false;

            using (var input = File.OpenRead(filePath))
            using (var zin = new ZipArchive(input, ZipArchiveMode.Read, false))
            {
                var ctEntry = zin.GetEntry("[Content_Types].xml");
                if (ctEntry != null)
                {
                    try
                    {
                        var xd = XDocument.Load(ctEntry.Open());
                        var types = xd.Root;
                        // check Default extension xml exists with application/xml
                        bool hasXmlDefault = false;
                        if (types != null)
                        {
                            foreach (var d in types.Elements())
                            {
                            // Default elements are in no namespace
                            if (d.Name.LocalName == "Default")
                            {
                                var ext = (string?)d.Attribute("Extension");
                                var ct = (string?)d.Attribute("ContentType");
                                if (ext == "xml" && ct == "application/xml") hasXmlDefault = true;
                            }
                            }
                        }

                        bool hasDocOverride = false, hasStylesOverride = false, hasNumberingOverride = false;
                        if (types != null)
                        {
                            foreach (var o in types.Elements())
                            {
                            if (o.Name.LocalName == "Override")
                            {
                                var part = (string?)o.Attribute("PartName");
                                if (part == "/word/document.xml") hasDocOverride = true;
                                if (part == "/word/styles.xml") hasStylesOverride = true;
                                if (part == "/word/numbering.xml") hasNumberingOverride = true;
                            }
                            }
                        }

                        if (!hasXmlDefault || !hasDocOverride || !hasStylesOverride || !hasNumberingOverride)
                        {
                            needsFix = true;
                        }
                    }
                    catch
                    {
                        needsFix = true;
                    }
                }
                // Validate relationships for leading slashes
                XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                var relRoot = zin.GetEntry("_rels/.rels");
                if (relRoot != null)
                {
                    try
                    {
                        var xd = XDocument.Load(relRoot.Open());
                        if (xd.Root != null)
                        {
                            foreach (var rel in xd.Root.Elements(relNs + "Relationship"))
                            {
                                var t = (string?)rel.Attribute("Target");
                                if (!string.IsNullOrEmpty(t) && t.StartsWith("/word/")) relsHaveLeadingSlash = true;
                            }
                        }
                    }
                    catch
                    {
                        relsHaveLeadingSlash = true;
                    }
                }

                var relDoc = zin.GetEntry("word/_rels/document.xml.rels");
                if (relDoc != null)
                {
                    try
                    {
                        var xd = XDocument.Load(relDoc.Open());
                        if (xd.Root != null)
                        {
                            foreach (var rel in xd.Root.Elements(relNs + "Relationship"))
                            {
                                var t = (string?)rel.Attribute("Target");
                                if (!string.IsNullOrEmpty(t) && t.StartsWith("/word/")) relsHaveLeadingSlash = true;
                            }
                        }
                    }
                    catch
                    {
                        relsHaveLeadingSlash = true;
                    }
                }

                // Validate that styles referenced in document.xml exist in styles.xml
                var docEntry = zin.GetEntry("word/document.xml");
                var stylesEntry = zin.GetEntry("word/styles.xml");
                var missingStyles = new HashSet<string>();
                try
                {
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    HashSet<string> declaredStyles = new();
                    if (stylesEntry != null)
                    {
                        var sx = XDocument.Load(stylesEntry.Open());
                        var styleEls = sx.Root?.Elements(w + "style");
                        if (styleEls != null)
                        {
                            foreach (var s in styleEls)
                            {
                                var id = (string?)s.Attribute("styleId");
                                if (!string.IsNullOrEmpty(id)) declaredStyles.Add(id);
                            }
                        }
                    }

                    if (docEntry != null)
                    {
                        var dx = XDocument.Load(docEntry.Open());
                        var pStyleEls = dx.Root?.Descendants(w + "pStyle");
                        if (pStyleEls != null)
                        {
                            foreach (var ps in pStyleEls)
                            {
                                var val = (string?)ps.Attribute("val");
                                if (!string.IsNullOrEmpty(val) && !declaredStyles.Contains(val)) missingStyles.Add(val);
                            }
                        }
                    }

                    if (missingStyles.Count > 0) needsFix = true;
                }
                catch
                {
                    // On any parse error, mark for fix to be safe
                    needsFix = true;
                }
            }

            if (!needsFix && !relsHaveLeadingSlash)
            {
                // Nothing to do
                return;
            }

            // Otherwise rewrite only the necessary parts
            var tempPath = filePath + ".tmp";
            using (var input = File.OpenRead(filePath))
            using (var zin = new ZipArchive(input, ZipArchiveMode.Read, false))
            using (var output = File.Create(tempPath))
            using (var zout = new ZipArchive(output, ZipArchiveMode.Create, false))
            {
                // If we need to add placeholder styles, prepare a map of current styles
                HashSet<string> existingStyles = new();
                XNamespace wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XDocument? stylesDoc = null;
                var stylesEntry = zin.GetEntry("word/styles.xml");
                if (stylesEntry != null)
                {
                    try
                    {
                        stylesDoc = XDocument.Load(stylesEntry.Open());
                        var styleEls = stylesDoc.Root?.Elements(wNs + "style");
                        if (styleEls != null)
                        {
                            foreach (var s in styleEls)
                            {
                                var id = (string?)s.Attribute("styleId");
                                if (!string.IsNullOrEmpty(id)) existingStyles.Add(id);
                            }
                        }
                    }
                    catch
                    {
                        stylesDoc = null;
                    }
                }

                foreach (var entry in zin.Entries)
                {
                    string name = entry.FullName;
                    if (name == "[Content_Types].xml")
                    {
                        // write canonical content types with correct main document type
                        string mainDocType = isDocument
                            ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                            : "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

                        var contentTypes =
                            "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                            "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                            "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                            $"  <Override PartName=\"/word/document.xml\" ContentType=\"{mainDocType}\"/>\n" +
                            "  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n" +
                            "  <Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>\n" +
                            "  <Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>\n" +
                            "</Types>\n";

                        var ze = zout.CreateEntry(name);
                        using var s = ze.Open();
                        using var wtr = new StreamWriter(s, Encoding.UTF8);
                        wtr.Write(contentTypes);
                    }
                    else if (name == "_rels/.rels" || name == "word/_rels/document.xml.rels")
                    {
                        // Parse and normalize relationship Targets to correct relative paths
                        using var r = entry.Open();
                        try
                        {
                            var xd = XDocument.Load(r);
                            XNamespace relNsLocal = "http://schemas.openxmlformats.org/package/2006/relationships";
                            if (xd.Root != null)
                            {
                                foreach (var rrel in xd.Root.Elements(relNsLocal + "Relationship"))
                                {
                                    var targ = (string?)rrel.Attribute("Target");
                                    if (string.IsNullOrEmpty(targ)) continue;

                                    // Remove any leading slash
                                    if (targ.StartsWith("/")) targ = targ.Substring(1);

                                    // For document.xml.rels (under word/), Targets should be relative to word/ root
                                    // So strip an extra leading "word/" if present
                                    if (name == "word/_rels/document.xml.rels" && targ.StartsWith("word/"))
                                    {
                                        targ = targ.Substring("word/".Length);
                                    }

                                    rrel.SetAttributeValue("Target", targ);
                                }
                            }
                            var ze = zout.CreateEntry(name);
                            using var s = ze.Open();
                            xd.Save(s);
                        }
                        catch
                        {
                            // fallback: copy as-is
                            var ze = zout.CreateEntry(name);
                            using var r2 = entry.Open();
                            using var w2 = ze.Open();
                            r2.CopyTo(w2);
                        }
                    }
                    else if (name == "word/styles.xml")
                    {
                        // write existing styles plus any missing placeholders
                        XDocument outStyles = stylesDoc ?? new XDocument(new XElement(wNs + "styles"));

                        // find referenced styles in document.xml
                        var docEntry = zin.GetEntry("word/document.xml");
                        HashSet<string> referenced = new();
                        if (docEntry != null)
                        {
                            try
                            {
                                var dx = XDocument.Load(docEntry.Open());
                                var pStyleEls = dx.Root?.Descendants(wNs + "pStyle");
                                if (pStyleEls != null)
                                {
                                    foreach (var ps in pStyleEls)
                                    {
                                        var val = (string?)ps.Attribute("val");
                                        if (!string.IsNullOrEmpty(val)) referenced.Add(val);
                                    }
                                }
                            }
                            catch { }
                        }

                        // Determine missing
                        var toAdd = referenced.Where(id => !existingStyles.Contains(id)).ToList();
                        if (toAdd.Count > 0)
                        {
                            var root = outStyles.Root ?? new XElement(wNs + "styles");
                            // ensure namespace declaration
                            if (root.Name.Namespace == XNamespace.None)
                            {
                                root = new XElement(wNs + "styles", root.Elements());
                            }
                            foreach (var id in toAdd)
                            {
                                var styleEl = new XElement(wNs + "style",
                                    new XAttribute("type", "paragraph"),
                                    new XAttribute("styleId", id),
                                    new XElement(wNs + "name", new XAttribute("val", id)),
                                    new XElement(wNs + "basedOn", new XAttribute("val", "Normal"))
                                );
                                root.Add(styleEl);
                            }
                            outStyles = new XDocument(root);
                        }

                        // Ensure styles are visible in Word's Styles pane: normalize and add missing elements
                        try
                        {
                            var root = outStyles.Root ?? new XElement(wNs + "styles");
                            // int pri = 100; // no longer used

                            // Ensure a latentStyles element exists with sane defaults (as in Word-created files)
                            var latent = root.Element(wNs + "latentStyles");
                            if (latent == null)
                            {
                                latent = new XElement(wNs + "latentStyles",
                                    new XAttribute(wNs + "defLockedState", "0"),
                                    new XAttribute(wNs + "defUIPriority", "99"),
                                    new XAttribute(wNs + "defSemiHidden", "0"),
                                    new XAttribute(wNs + "defUnhideWhenUsed", "0"),
                                    new XAttribute(wNs + "defQFormat", "0"),
                                    new XAttribute(wNs + "count", "0")
                                );
                                var docDefaults = root.Element(wNs + "docDefaults");
                                if (docDefaults != null)
                                {
                                    docDefaults.AddAfterSelf(latent);
                                }
                                else
                                {
                                    root.AddFirst(latent);
                                }
                            }
                            foreach (var s in root.Elements(wNs + "style"))
                            {
                                // ensure there is a name element; if missing, add it first
                                var nameEl = s.Element(wNs + "name");
                                if (nameEl == null)
                                {
                                    // use styleId as name fallback
                                    var sid = (string?)s.Attribute("styleId") ?? "";
                                    nameEl = new XElement(wNs + "name", new XAttribute("val", sid));
                                    s.AddFirst(nameEl);
                                }

                                // Add qFormat if missing (required for custom styles to be visible)
                                if (s.Element(wNs + "qFormat") == null)
                                {
                                    nameEl.AddAfterSelf(new XElement(wNs + "qFormat"));
                                }

                                // Remove unhideWhenUsed and semiHidden if present (not needed, may interfere)
                                var semiHiddenEl = s.Element(wNs + "semiHidden");
                                semiHiddenEl?.Remove();

                                var unhideEl = s.Element(wNs + "unhideWhenUsed");
                                unhideEl?.Remove();                                // Fix customStyle attribute: Word expects "1" not "true"
                                var customStyleAttr = s.Attribute(wNs + "customStyle");
                                if (customStyleAttr != null && customStyleAttr.Value == "true")
                                {
                                    customStyleAttr.Value = "1";
                                }

                                // Fix attribute order: customStyle should come BEFORE styleId (as in demo.docx)
                                // OpenXML SDK writes them alphabetically, but Word may care about order
                                if (customStyleAttr != null)
                                {
                                    var styleIdAttr = s.Attribute(wNs + "styleId");
                                    var typeAttr = s.Attribute(wNs + "type");
                                    var defaultAttr = s.Attribute(wNs + "default");

                                    // Save values
                                    string styleIdValue = styleIdAttr?.Value ?? "";
                                    string customStyleValue = customStyleAttr.Value;
                                    string? defaultValue = defaultAttr?.Value;

                                    // Remove all and re-add in correct order: type, customStyle, styleId, default
                                    if (styleIdAttr != null && typeAttr != null)
                                    {
                                        customStyleAttr.Remove();
                                        styleIdAttr.Remove();
                                        defaultAttr?.Remove();

                                        // Add attributes in the order Word expects (type already present)
                                        s.Add(new XAttribute(wNs + "customStyle", customStyleValue));
                                        s.Add(new XAttribute(wNs + "styleId", styleIdValue));
                                        if (defaultValue != null)
                                        {
                                            s.Add(new XAttribute(wNs + "default", defaultValue));
                                        }
                                    }
                                }

                                // Refresh the customStyleAttr reference after potential re-creation
                                customStyleAttr = s.Attribute(wNs + "customStyle");

                                // Add rsid if missing (Word tracking ID) - custom styles should have this
                                if (customStyleAttr != null && s.Element(wNs + "rsid") == null)
                                {
                                    var after3 = s.Element(wNs + "qFormat") ?? s.Element(wNs + "uiPriority") ?? nameEl;
                                    after3.AddAfterSelf(new XElement(wNs + "rsid", new XAttribute(wNs + "val", "00000001")));
                                }
                            }
                            outStyles = new XDocument(root);
                        }
                        catch { }

                        var ze = zout.CreateEntry(name);
                        using var outStream = ze.Open();
                        outStyles.Save(outStream);
                    }
                    else if (name == "word/_rels/document.xml.rels")
                    {
                        // Ensure settings relationship exists
                        using var r = entry.Open();
                        try
                        {
                            var xd = XDocument.Load(r);
                            XNamespace relNsLocal = "http://schemas.openxmlformats.org/package/2006/relationships";
                            var rootR = xd.Root ?? new XElement(relNsLocal + "Relationships");
                            bool hasSettingsRel = rootR.Elements(relNsLocal + "Relationship")
                                .Any(e => ((string?)e.Attribute("Type")) == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
                            if (!hasSettingsRel)
                            {
                                rootR.Add(new XElement(relNsLocal + "Relationship",
                                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"),
                                    new XAttribute("Target", "settings.xml"),
                                    new XAttribute("Id", "Rsettings")
                                ));
                            }
                            var ze = zout.CreateEntry(name);
                            using var s = ze.Open();
                            xd.Save(s);
                        }
                        catch
                        {
                            var ze = zout.CreateEntry(name);
                            using var r2 = entry.Open();
                            using var w2 = ze.Open();
                            r2.CopyTo(w2);
                        }
                    }
                    else if (name == "word/settings.xml")
                    {
                        // Copy existing settings if present
                        var ze = zout.CreateEntry(name);
                        using var r = entry.Open();
                        using var w = ze.Open();
                        r.CopyTo(w);
                    }
                    else
                    {
                        var ze = zout.CreateEntry(name);
                        using var r = entry.Open();
                        using var w = ze.Open();
                        r.CopyTo(w);
                    }
                }

                // If settings.xml was missing, add a minimal settings part to help Word UI
                bool hasSettingsPart = zin.GetEntry("word/settings.xml") != null;
                if (!hasSettingsPart)
                {
                    var ze = zout.CreateEntry("word/settings.xml");
                    using var s = ze.Open();
                    var settingsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:stylePaneFormatFilter w:allStyles=\"1\" w:visibleStyles=\"1\"/>" +
                    "<w:themeFontLang w:val=\"en-US\"/>" +
                    "</w:settings>";
                    var wtr = new StreamWriter(s, Encoding.UTF8);
                    wtr.Write(settingsXml);
                    wtr.Flush();
                }
            }

            // Replace original with fixed file
            File.Replace(tempPath, filePath, null);
        }

    }
}
