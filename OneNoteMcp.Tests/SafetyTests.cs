// Safety-fix regression tests: XML attribute escaping for model-supplied IDs
// and magic-byte validation for embedded images.

using OneNoteMcp;
using Xunit;

namespace OneNoteMcp.Tests;

public class SafetyTests
{
    [Fact]
    public void XmlAttrEscapesAttributeBreakoutCharacters()
    {
        var hostile = "x\" /><one:Notebook ID=\"stolen\"/><one:Section ID=\"x";
        var escaped = OneNoteCom.XmlAttr(hostile);
        Assert.DoesNotContain("\"", escaped);
        Assert.DoesNotContain("<", escaped);
        Assert.DoesNotContain(">", escaped);
    }

    [Fact]
    public void XmlAttrPassesRealObjectIdsThrough()
    {
        var id = "{9DF45242-9B04-46BE-9892-7FB0C1A19801}{1}{B0}";
        Assert.Equal(id, OneNoteCom.XmlAttr(id));
    }

    [Fact]
    public void XmlAttrKeepsEscapedIdWellFormedInsideAnAttribute()
    {
        var hostile = "x\"><one:Page ID=\"y";
        var xml = $"<one:Section xmlns:one=\"{OneNoteCom.OneNS}\" ID=\"{OneNoteCom.XmlAttr(hostile)}\" />";
        var doc = new System.Xml.XmlDocument();
        doc.LoadXml(xml);
        Assert.Equal(hostile, doc.DocumentElement!.GetAttribute("ID"));
    }

    [Theory]
    [InlineData(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A })] // PNG
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46 })] // JPEG
    [InlineData(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00 })] // GIF89a
    [InlineData(new byte[] { 0x42, 0x4D, 0x36, 0x00, 0x00, 0x00, 0x00, 0x00 })] // BMP
    [InlineData(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00 })] // TIFF LE
    [InlineData(new byte[] { 0x4D, 0x4D, 0x00, 0x2A, 0x00, 0x00, 0x00, 0x08 })] // TIFF BE
    public void AcceptsRealImageHeaders(byte[] header) =>
        Assert.True(OneNoteCom.LooksLikeSupportedImage(header));

    [Fact]
    public void ExportTargetAcceptsAbsolutePathWithMatchingExtension()
    {
        var dir = Path.GetTempPath();
        var target = Path.Combine(dir, "note.pdf");
        Assert.Equal(Path.GetFullPath(target), OneNoteCom.ValidateExportTarget(target, "pdf"));
    }

    [Theory]
    [InlineData("relative\\note.pdf", "pdf")]     // not absolute
    [InlineData("note.pdf", "pdf")]               // not absolute
    public void ExportTargetRejectsRelativePaths(string target, string fmt) =>
        Assert.Throws<ArgumentException>(() => OneNoteCom.ValidateExportTarget(target, fmt));

    [Theory]
    [InlineData("evil.bat", "pdf")]
    [InlineData("evil.exe", "html")]
    [InlineData("evil.ps1", "docx")]
    public void ExportTargetRejectsMismatchedExtensions(string fileName, string fmt)
    {
        var target = Path.Combine(Path.GetTempPath(), fileName);
        Assert.Throws<ArgumentException>(() => OneNoteCom.ValidateExportTarget(target, fmt));
    }

    [Fact]
    public void ExportTargetRejectsMissingDirectory()
    {
        var target = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "note.pdf");
        Assert.Throws<ArgumentException>(() => OneNoteCom.ValidateExportTarget(target, "pdf"));
    }

    [Theory]
    [InlineData("Team Notes", "")]
    [InlineData("Team Notes", "C:\\Notes")]
    public void NotebookLocationAcceptsPlainNames(string name, string path) =>
        OneNoteCom.ValidateNotebookLocation(name, path);

    [Theory]
    [InlineData("..", "")]                                    // traversal
    [InlineData("..\\..\\Startup", "")]                       // traversal
    [InlineData("C:\\Users\\victim\\evil", "")]               // rooted name overrides folder
    [InlineData("sub\\dir", "")]                              // separator smuggling
    [InlineData("x", "relative\\folder")]                     // non-absolute path
    public void NotebookLocationRejectsEscapes(string name, string path) =>
        Assert.Throws<ArgumentException>(() => OneNoteCom.ValidateNotebookLocation(name, path));

    [Theory]
    [InlineData("-----BEGIN OPENSSH PRIVATE KEY-----")]
    [InlineData("SQLite format 3\0")]
    [InlineData("MZ\x90\0\x03\0\0\0")] // PE executable
    [InlineData("{\"credentials\":1}")]
    [InlineData("")]
    [InlineData("short")]
    public void RejectsNonImageContent(string content) =>
        Assert.False(OneNoteCom.LooksLikeSupportedImage(System.Text.Encoding.Latin1.GetBytes(content)));
}
