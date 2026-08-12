// Direct OneNote COM layer. Replaces the PowerShell bridge the TypeScript
// version needed (Node cannot speak COM).
//
// Early-bound vtable dispatch, NOT IDispatch: with x64 Click-to-Run Office the
// OneNote typelib is registered only under the Win32 key, so every IDispatch
// metadata path fails from a 64-bit .NET process — GetTypeInfo returns E_FAIL,
// GetIDsOfNames/Invoke return TYPE_E_LIBNOTREGISTERED (this is the same wall
// pywin32 hits, and C# `dynamic` hits it too). Calling straight through the
// IApplication dual-interface vtable (IID 452AC71A-B655-4967-A208-A4CC39DD7949,
// slot layout below matches the OneNote 15.0 typelib) bypasses all typelib
// lookups and works.
//
// All COM calls are marshalled onto one dedicated STA thread: the OneNote COM
// API wants a single-threaded apartment, and a single work queue also
// serializes calls the way the persistent bridge process did.

using System.Collections.Concurrent;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OneNoteMcp;

// OneNote 15.0 IApplication dual interface. Method order MUST match the
// typelib vtable exactly (slots 7+ after IUnknown/IDispatch); methods this
// server never calls are still declared to keep every slot aligned.
// Optional IDL parameters must all be passed explicitly on a vtable call.
[ComImport, Guid("452AC71A-B655-4967-A208-A4CC39DD7949"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
internal interface IOneNoteApplication
{
    void GetHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID, int hsScope,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut, int xsSchema);
    void UpdateHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrChangesXmlIn, int xsSchema);
    void OpenHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrPath,
        [MarshalAs(UnmanagedType.BStr)] string bstrRelativeToObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrObjectID, int cftIfNotExist);
    void DeleteHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        double dateExpectedLastModified, [MarshalAs(UnmanagedType.VariantBool)] bool deletePermanently);
    void CreateNewPage([MarshalAs(UnmanagedType.BStr)] string bstrSectionID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrPageID, int npsNewPageStyle);
    void CloseNotebook([MarshalAs(UnmanagedType.BStr)] string bstrNotebookID,
        [MarshalAs(UnmanagedType.VariantBool)] bool force);
    void GetHierarchyParent([MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrParentID);
    void GetPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrPageXmlOut, int pageInfoToExport, int xsSchema);
    void UpdatePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageChangesXmlIn,
        double dateExpectedLastModified, int xsSchema, [MarshalAs(UnmanagedType.VariantBool)] bool force);
    void GetBinaryPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [MarshalAs(UnmanagedType.BStr)] string bstrCallbackID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrBinaryObjectB64Out);
    void DeletePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID,
        [MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        double dateExpectedLastModified, [MarshalAs(UnmanagedType.VariantBool)] bool force);
    void NavigateTo([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyObjectID,
        [MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        [MarshalAs(UnmanagedType.VariantBool)] bool fNewWindow);
    void NavigateToUrl([MarshalAs(UnmanagedType.BStr)] string bstrUrl,
        [MarshalAs(UnmanagedType.VariantBool)] bool fNewWindow);
    void Publish([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID,
        [MarshalAs(UnmanagedType.BStr)] string bstrTargetFilePath, int pfPublishFormat,
        [MarshalAs(UnmanagedType.BStr)] string bstrCLSIDofExporter);
    void OpenPackage([MarshalAs(UnmanagedType.BStr)] string bstrPathPackage,
        [MarshalAs(UnmanagedType.BStr)] string bstrPathDest,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrPathOut);
    void GetHyperlinkToObject([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID,
        [MarshalAs(UnmanagedType.BStr)] string bstrPageContentObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHyperlinkOut);
    void FindPages([MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID,
        [MarshalAs(UnmanagedType.BStr)] string bstrSearchString,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut,
        [MarshalAs(UnmanagedType.VariantBool)] bool fIncludeUnindexedPages,
        [MarshalAs(UnmanagedType.VariantBool)] bool fDisplay, int xsSchema);
    void FindMeta([MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID,
        [MarshalAs(UnmanagedType.BStr)] string bstrSearchStringName,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut,
        [MarshalAs(UnmanagedType.VariantBool)] bool fIncludeUnindexedPages, int xsSchema);
    void GetSpecialLocation(int slToGet,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrSpecialLocationPath);
}

internal sealed class StaWorker : IDisposable
{
    private readonly BlockingCollection<Action> _queue = new();
    private readonly Thread _thread;

    public StaWorker()
    {
        _thread = new Thread(() =>
        {
            foreach (var work in _queue.GetConsumingEnumerable()) work();
        })
        {
            IsBackground = true,
            Name = "OneNoteComSta",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public T Run<T>(Func<T> fn)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try { tcs.SetResult(fn()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task.GetAwaiter().GetResult();
    }

    public void Dispose() => _queue.CompleteAdding();
}

public sealed class OneNoteCom : IDisposable
{
    public const string OneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private static readonly Dictionary<string, int> ScopeMap = new()
    {
        ["notebooks"] = 2,
        ["sections"] = 3,
        ["pages"] = 4,
    };

    private const int CftNotebook = 1;
    private const int CftFolder = 2; // section group
    private const int CftSection = 3;
    private const int SlDefaultNotebookFolder = 2;
    private const int Xs2013 = 2;   // XMLSchema.xs2013 — the schema the PS bridge got by default
    private const int PiBasic = 0;  // PageInfo.piBasic — text content, no binary data

    private static readonly Dictionary<string, int> PubFormat = new()
    {
        ["onenote"] = 0, ["package"] = 1, ["mhtml"] = 2, ["pdf"] = 3,
        ["xps"] = 4, ["word"] = 5, ["docx"] = 5, ["emf"] = 6, ["html"] = 7,
    };

    // COM errors that mean "the OneNote we bound to is gone" — recreate the
    // object and retry once rather than surfacing a dead-handle error.
    private static readonly Regex DisconnectPatterns = new(
        "RPC server is unavailable|The object invoked has disconnected|been severed|0x800706BA|0x80010108|0x800706BE",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly int[] DisconnectHResults =
    {
        unchecked((int)0x800706BA), // RPC_S_SERVER_UNAVAILABLE
        unchecked((int)0x80010108), // RPC_E_DISCONNECTED
        unchecked((int)0x800706BE), // RPC_S_CALL_FAILED
    };

    private readonly StaWorker _worker = new();
    private object? _app;

    private IOneNoteApplication App => (IOneNoteApplication)(_app ?? throw new InvalidOperationException("OneNote COM object not initialized"));

    private void EnsureApp()
    {
        if (_app is not null) return;
        var type = Type.GetTypeFromProgID("OneNote.Application")
            ?? throw new InvalidOperationException(
                "OneNote.Application is not registered. Install the OneNote desktop app from Office.");
        _app = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Could not create the OneNote.Application COM object.");
    }

    private static bool IsDisconnect(Exception ex) =>
        (ex is COMException com && DisconnectHResults.Contains(com.HResult))
        || DisconnectPatterns.IsMatch(ex.Message);

    private T Invoke<T>(Func<T> fn) => _worker.Run(() =>
    {
        EnsureApp();
        try
        {
            return fn();
        }
        catch (Exception ex) when (IsDisconnect(ex))
        {
            // OneNote was closed/restarted; rebind to a live instance and retry once.
            if (_app is not null) { Marshal.FinalReleaseComObject(_app); _app = null; }
            EnsureApp();
            return fn();
        }
    });

    public void Dispose()
    {
        _worker.Run<object?>(() =>
        {
            if (_app is not null) { Marshal.FinalReleaseComObject(_app); _app = null; }
            return null;
        });
        _worker.Dispose();
    }

    // --- XML helpers ---------------------------------------------------------

    private static XmlDocument Parse(string xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);
        return doc;
    }

    // Object IDs are model-supplied (and can originate from untrusted note
    // content), so any ID interpolated into hand-built XML must be escaped or
    // a crafted "ID" breaks out of the attribute and injects hierarchy XML.
    internal static string XmlAttr(string value) =>
        System.Security.SecurityElement.Escape(value);

    private static string StripHtml(string? fragment)
    {
        if (fragment is null) return "";
        var text = Regex.Replace(fragment, "<br[^>]*>", "\n");
        text = Regex.Replace(text, "<[^>]+>", "");
        return WebUtility.HtmlDecode(text);
    }

    private static HierarchyNode NodeSummary(XmlElement el)
    {
        var name = el.GetAttribute("name");
        if (name.Length == 0) name = el.GetAttribute("nickname");
        var node = new HierarchyNode { Kind = el.LocalName, Id = el.GetAttribute("ID"), Name = name };
        if (el.GetAttribute("isCurrentlyViewed") == "true") node.CurrentlyViewed = true;
        var lastModified = el.GetAttribute("lastModifiedTime");
        if (lastModified.Length > 0) node.LastModifiedTime = lastModified;
        var dateTime = el.GetAttribute("dateTime");
        if (dateTime.Length > 0) node.DateTime = dateTime;
        return node;
    }

    private static readonly string[] HierarchyKinds = { "Notebook", "SectionGroup", "Section", "Page" };

    private static HierarchyNode HierarchyTree(XmlElement el)
    {
        var node = NodeSummary(el);
        var children = new List<HierarchyNode>();
        foreach (XmlNode child in el.ChildNodes)
        {
            if (child is XmlElement childEl && HierarchyKinds.Contains(childEl.LocalName))
                children.Add(HierarchyTree(childEl));
        }
        if (children.Count > 0) node.Children = children;
        return node;
    }

    // --- Read ops ------------------------------------------------------------

    public List<HierarchyNode> Hierarchy(string scope, string startId) => Invoke(() =>
    {
        var scopeValue = ScopeMap.TryGetValue(scope.ToLowerInvariant(), out var s) ? s : 4;
        string xmlOut = "";
        App.GetHierarchy(startId, scopeValue, out xmlOut, Xs2013);
        var doc = Parse(xmlOut);
        var items = new List<HierarchyNode>();
        // Scope "notebooks" returns the Notebooks root whose children are the
        // notebooks; other scopes nest the same way, so children of the
        // document element are always the items the caller asked about.
        foreach (XmlNode el in doc.DocumentElement!.ChildNodes)
        {
            if (el is XmlElement element) items.Add(HierarchyTree(element));
        }
        return items;
    });

    public PageRecord GetPage(string pageId) => Invoke(() =>
    {
        string xmlOut = "";
        App.GetPageContent(pageId, out xmlOut, PiBasic, Xs2013);
        var doc = Parse(xmlOut);
        var title = "";
        var titleEls = doc.GetElementsByTagName("Title", OneNS);
        if (titleEls.Count > 0 && titleEls[0] is XmlElement titleEl)
        {
            var titleT = titleEl.GetElementsByTagName("T", OneNS);
            if (titleT.Count > 0) title = StripHtml(titleT[0]!.InnerText);
        }
        var lines = new List<string>();
        foreach (XmlNode outline in doc.GetElementsByTagName("Outline", OneNS))
        {
            foreach (XmlNode t in ((XmlElement)outline).GetElementsByTagName("T", OneNS))
                lines.Add(StripHtml(t.InnerText));
        }
        return new PageRecord { PageId = pageId, Title = title, Text = string.Join("\n", lines) };
    });

    public Dictionary<string, object?> Search(string query, string startId) => Invoke(() =>
    {
        string xmlOut = "";
        App.FindPages(startId, query, out xmlOut, false, false, Xs2013);
        var doc = Parse(xmlOut);
        var pages = new List<Dictionary<string, object?>>();
        foreach (XmlNode page in doc.GetElementsByTagName("Page", OneNS))
        {
            var el = (XmlElement)page;
            pages.Add(new Dictionary<string, object?> { ["id"] = el.GetAttribute("ID"), ["name"] = el.GetAttribute("name") });
        }
        return new Dictionary<string, object?> { ["query"] = query, ["pages"] = pages };
    });

    // --- Page content ops ----------------------------------------------------

    private int AppendTextToPage(string pageId, string text)
    {
        string xmlOut = "";
        App.GetPageContent(pageId, out xmlOut, PiBasic, Xs2013);
        var doc = Parse(xmlOut);
        var outline = doc.CreateElement("one", "Outline", OneNS);
        var children = doc.CreateElement("one", "OEChildren", OneNS);
        outline.AppendChild(children);
        var paragraphs = text.Split('\n');
        foreach (var para in paragraphs)
        {
            var oe = doc.CreateElement("one", "OE", OneNS);
            var t = doc.CreateElement("one", "T", OneNS);
            // OneNote interprets T CDATA as HTML, so plain text must be
            // HTML-encoded or literal "<task>"-style angle brackets get parsed
            // as tags and vanish.
            t.AppendChild(doc.CreateCDataSection(WebUtility.HtmlEncode(para)));
            oe.AppendChild(t);
            children.AppendChild(oe);
        }
        doc.DocumentElement!.AppendChild(outline);
        App.UpdatePageContent(doc.OuterXml, 0.0, Xs2013, false);
        return paragraphs.Length;
    }

    public Dictionary<string, object?> AppendPage(string pageId, string text) => Invoke(() =>
        new Dictionary<string, object?> { ["appended_paragraphs"] = AppendTextToPage(pageId, text) });

    private void SetPageTitleCore(string pageId, string title)
    {
        string xmlOut = "";
        App.GetPageContent(pageId, out xmlOut, PiBasic, Xs2013);
        var doc = Parse(xmlOut);
        XmlElement titleEl;
        var titleEls = doc.GetElementsByTagName("Title", OneNS);
        if (titleEls.Count > 0)
        {
            titleEl = (XmlElement)titleEls[0]!;
        }
        else
        {
            titleEl = doc.CreateElement("one", "Title", OneNS);
            doc.DocumentElement!.PrependChild(titleEl);
        }
        XmlElement t;
        var tEls = titleEl.GetElementsByTagName("T", OneNS);
        if (tEls.Count > 0)
        {
            t = (XmlElement)tEls[0]!;
        }
        else
        {
            var oe = doc.CreateElement("one", "OE", OneNS);
            t = doc.CreateElement("one", "T", OneNS);
            oe.AppendChild(t);
            titleEl.AppendChild(oe);
        }
        t.RemoveAll();
        t.AppendChild(doc.CreateCDataSection(WebUtility.HtmlEncode(title)));
        App.UpdatePageContent(doc.OuterXml, 0.0, Xs2013, false);
    }

    public Dictionary<string, object?> CreatePage(string sectionId, string title, string body) => Invoke(() =>
    {
        string pageId = "";
        App.CreateNewPage(sectionId, out pageId, 0);
        if (title.Length > 0) SetPageTitleCore(pageId, title);
        if (body.Length > 0) AppendTextToPage(pageId, body);
        return new Dictionary<string, object?> { ["page_id"] = pageId };
    });

    public Dictionary<string, object?> RenamePage(string pageId, string newTitle) => Invoke(() =>
    {
        SetPageTitleCore(pageId, newTitle);
        return new Dictionary<string, object?> { ["page_id"] = pageId, ["title"] = newTitle };
    });

    public Dictionary<string, object?> UpdatePage(string pageId, string content, string mode) => Invoke(() =>
    {
        var normalizedMode = mode.ToLowerInvariant();
        if (normalizedMode == "append")
        {
            var appended = AppendTextToPage(pageId, content);
            return new Dictionary<string, object?> { ["page_id"] = pageId, ["mode"] = "append", ["paragraphs"] = appended };
        }
        // replace: remove existing Outline objects, then append fresh content.
        string xmlOut = "";
        App.GetPageContent(pageId, out xmlOut, PiBasic, Xs2013);
        var doc = Parse(xmlOut);
        var outlineIds = new List<string>();
        foreach (XmlNode outline in doc.GetElementsByTagName("Outline", OneNS))
        {
            var oid = ((XmlElement)outline).GetAttribute("objectID");
            if (oid.Length > 0) outlineIds.Add(oid);
        }
        foreach (var oid in outlineIds)
        {
            try { App.DeletePageContent(pageId, oid, 0.0, false); } catch { /* stale outline id: nothing to remove */ }
        }
        var count = AppendTextToPage(pageId, content);
        return new Dictionary<string, object?> { ["page_id"] = pageId, ["mode"] = "replace", ["paragraphs"] = count };
    });

    // --- Navigation / delete -------------------------------------------------

    public Dictionary<string, object?> Navigate(string objectId) => Invoke(() =>
    {
        App.NavigateTo(objectId, "", false);
        return new Dictionary<string, object?> { ["navigated"] = objectId };
    });

    private void RemoveObject(string objectId, bool permanent)
    {
        // deletePermanently=false moves the object to the notebook's recycle
        // bin (recoverable) rather than erasing it. Works for any hierarchy
        // object — page, section, section group, or notebook. DATE 0 for
        // dateExpectedLastModified means "no concurrency check".
        App.DeleteHierarchy(objectId, 0.0, permanent);
    }

    public Dictionary<string, object?> DeleteObject(string objectId, bool permanent) => Invoke(() =>
    {
        RemoveObject(objectId, permanent);
        return new Dictionary<string, object?> { ["deleted"] = objectId, ["permanent"] = permanent };
    });

    // --- Hierarchy create / rename / move / reorder --------------------------

    private XmlElement? FindHierNode(string startId, int scope, string localName, string targetId)
    {
        string xmlOut = "";
        App.GetHierarchy(startId, scope, out xmlOut, Xs2013);
        var doc = Parse(xmlOut);
        foreach (XmlNode node in doc.GetElementsByTagName(localName, OneNS))
        {
            if (node is XmlElement el && el.GetAttribute("ID") == targetId) return el;
        }
        return null;
    }

    public Dictionary<string, object?> CreateSection(string parentId, string sectionName) => Invoke(() =>
    {
        // parentId may be a notebook OR a section group — OpenHierarchy
        // resolves the section path relative to it.
        var leaf = sectionName;
        if (!leaf.EndsWith(".one", StringComparison.OrdinalIgnoreCase)) leaf += ".one";
        string id = "";
        App.OpenHierarchy(leaf, parentId, out id, CftSection);
        return new Dictionary<string, object?> { ["section_id"] = id, ["name"] = sectionName };
    });

    public Dictionary<string, object?> CreateSectionGroup(string parentId, string name) => Invoke(() =>
    {
        string id = "";
        App.OpenHierarchy(name, parentId, out id, CftFolder);
        return new Dictionary<string, object?> { ["section_group_id"] = id, ["name"] = name };
    });

    // Path.Combine silently discards the folder when the second argument is
    // rooted, so an unvalidated name (or a name with separators / "..") lets a
    // notebook be created at an arbitrary filesystem location instead of
    // inside the chosen folder.
    internal static void ValidateNotebookLocation(string name, string path)
    {
        if (name.Length == 0 || name is "." or ".."
            || name.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0)
            throw new ArgumentException($"name must be a plain folder name (no path separators): {name}");
        if (path.Length > 0 && !System.IO.Path.IsPathFullyQualified(path))
            throw new ArgumentException($"path must be an absolute folder: {path}");
    }

    public Dictionary<string, object?> CreateNotebook(string name, string path) => Invoke(() =>
    {
        ValidateNotebookLocation(name, path);
        var folder = path;
        if (folder.Length == 0)
        {
            string loc = "";
            App.GetSpecialLocation(SlDefaultNotebookFolder, out loc);
            folder = loc;
        }
        var full = System.IO.Path.GetFullPath(System.IO.Path.Combine(folder, name));
        string id = "";
        App.OpenHierarchy(full, "", out id, CftNotebook);
        return new Dictionary<string, object?> { ["notebook_id"] = id, ["name"] = name, ["path"] = full };
    });

    public Dictionary<string, object?> RenameSection(string sectionId, string newName) => Invoke(() =>
    {
        // Scan sections scope so the section node is found regardless of nesting.
        var sec = FindHierNode("", 3, "Section", sectionId)
            ?? throw new InvalidOperationException($"section not found: {sectionId}");
        sec.SetAttribute("name", newName);
        App.UpdateHierarchy(sec.OuterXml, Xs2013);
        return new Dictionary<string, object?> { ["section_id"] = sectionId, ["name"] = newName };
    });

    private HashSet<string> GetSectionPageIds(string sectionId)
    {
        string xmlOut = "";
        App.GetHierarchy(sectionId, 4, out xmlOut, Xs2013);
        var doc = Parse(xmlOut);
        var ids = new HashSet<string>();
        foreach (XmlNode page in doc.GetElementsByTagName("Page", OneNS))
            ids.Add(((XmlElement)page).GetAttribute("ID"));
        return ids;
    }

    public Dictionary<string, object?> MovePage(string pageId, string targetSectionId) => Invoke(() =>
    {
        // OneNote regenerates a moved page's entire object ID (both the section
        // GUID and the page's own GUID change). Snapshot the target section's
        // page IDs before and after so we can hand back the new ID as the set
        // difference.
        var before = GetSectionPageIds(targetSectionId);
        var xml =
            $"<one:Section xmlns:one=\"{OneNS}\" ID=\"{XmlAttr(targetSectionId)}\">" +
            $"<one:Page ID=\"{XmlAttr(pageId)}\" /></one:Section>";
        App.UpdateHierarchy(xml, Xs2013);
        var after = GetSectionPageIds(targetSectionId);
        string? newId = after.FirstOrDefault(k => !before.Contains(k));
        return new Dictionary<string, object?>
        {
            ["page_id"] = newId ?? pageId,
            ["previous_page_id"] = pageId,
            ["section_id"] = targetSectionId,
            ["id_changed"] = newId is not null,
        };
    });

    public Dictionary<string, object?> MoveSection(string sectionId, string targetParentId) => Invoke(() =>
    {
        // The target parent is a notebook or a section group; reparent by
        // submitting it with the section as a child. Try SectionGroup wrapper
        // first, then Notebook, so either parent kind works.
        var secXml = $"<one:Section ID=\"{XmlAttr(sectionId)}\" />";
        var attempts = new[]
        {
            $"<one:SectionGroup xmlns:one=\"{OneNS}\" ID=\"{XmlAttr(targetParentId)}\">{secXml}</one:SectionGroup>",
            $"<one:Notebook xmlns:one=\"{OneNS}\" ID=\"{XmlAttr(targetParentId)}\">{secXml}</one:Notebook>",
        };
        string? lastErr = null;
        foreach (var xml in attempts)
        {
            try
            {
                App.UpdateHierarchy(xml, Xs2013);
                return new Dictionary<string, object?> { ["section_id"] = sectionId, ["parent_id"] = targetParentId };
            }
            catch (Exception ex)
            {
                lastErr = ex.Message;
                if (IsDisconnect(ex)) throw;
            }
        }
        throw new InvalidOperationException($"move_section failed: {lastErr}");
    });

    private void ReorderChildren(string parentId, int scope, string childLocal, string moveId, string beforeId, string afterId)
    {
        // Fetch the parent, reorder its child elements of type childLocal so
        // moveId lands before/after the reference, then resubmit the parent.
        string xmlOut = "";
        App.GetHierarchy(parentId, scope, out xmlOut, Xs2013);
        var doc = Parse(xmlOut);
        XmlElement? parent = null;
        if (parentId.Length > 0)
        {
            foreach (XmlNode node in doc.GetElementsByTagName("*"))
            {
                if (node is XmlElement el && el.GetAttribute("ID") == parentId) { parent = el; break; }
            }
        }
        parent ??= doc.DocumentElement!;
        var kids = new List<XmlElement>();
        foreach (XmlNode c in parent.ChildNodes)
        {
            if (c is XmlElement el && el.LocalName == childLocal) kids.Add(el);
        }
        var moveNode = kids.FirstOrDefault(k => k.GetAttribute("ID") == moveId)
            ?? throw new InvalidOperationException($"{childLocal} not found in parent: {moveId}");
        var refId = beforeId.Length > 0 ? beforeId : afterId;
        var refNode = kids.FirstOrDefault(k => k.GetAttribute("ID") == refId)
            ?? throw new InvalidOperationException($"reference {childLocal} not found: {refId}");
        parent.RemoveChild(moveNode);
        if (beforeId.Length > 0) parent.InsertBefore(moveNode, refNode);
        else parent.InsertAfter(moveNode, refNode);
        App.UpdateHierarchy(parent.OuterXml, Xs2013);
    }

    public Dictionary<string, object?> ReorderPages(string sectionId, string pageId, string beforePageId, string afterPageId) => Invoke(() =>
    {
        if (beforePageId.Length == 0 && afterPageId.Length == 0)
            throw new ArgumentException("before_page_id or after_page_id is required");
        ReorderChildren(sectionId, 4, "Page", pageId, beforePageId, afterPageId);
        return new Dictionary<string, object?> { ["page_id"] = pageId, ["reordered"] = true };
    });

    public Dictionary<string, object?> ReorderSections(string parentId, string sectionId, string beforeSectionId, string afterSectionId) => Invoke(() =>
    {
        if (beforeSectionId.Length == 0 && afterSectionId.Length == 0)
            throw new ArgumentException("before_section_id or after_section_id is required");
        ReorderChildren(parentId, 3, "Section", sectionId, beforeSectionId, afterSectionId);
        return new Dictionary<string, object?> { ["section_id"] = sectionId, ["reordered"] = true };
    });

    // --- HTML -> OneNote XML translation -------------------------------------
    //
    // OneNote's UpdatePageContent only accepts *inline* HTML (span/a) inside a
    // one:T CDATA. Block structure — headings, paragraphs, lists, tables —
    // must be expressed as one:OE / one:List / one:Table elements, or the
    // whole update is rejected with 0x80042009 (invalid XML). These helpers
    // walk a well-formed XHTML fragment and emit the corresponding OneNote XML.

    private static readonly (string Entity, string Numeric)[] EntityMap =
    {
        ("&nbsp;", "&#160;"), ("&mdash;", "&#8212;"), ("&ndash;", "&#8211;"),
        ("&bull;", "&#8226;"), ("&hellip;", "&#8230;"), ("&middot;", "&#183;"),
        ("&lsquo;", "&#8216;"), ("&rsquo;", "&#8217;"), ("&ldquo;", "&#8220;"),
        ("&rdquo;", "&#8221;"), ("&copy;", "&#169;"), ("&reg;", "&#174;"),
        ("&trade;", "&#8482;"), ("&rarr;", "&#8594;"), ("&larr;", "&#8592;"),
    };

    private static string PrepareXhtmlFragment(string html)
    {
        // XML parsing knows only the five XML entities; map the common HTML
        // ones to numeric references so fragments with &nbsp; etc. still parse.
        var s = html;
        foreach (var (entity, numeric) in EntityMap) s = s.Replace(entity, numeric);
        return s;
    }

    private static string ProtectCdata(string s) =>
        // "]]>" inside CDATA would terminate the section; OneNote decodes the
        // HTML entity form back to the literal characters when rendering.
        s.Replace("]]>", "]]&gt;");

    private static string ConvertInlineHtml(XmlNode node)
    {
        // Flatten an element's children into the inline HTML subset OneNote
        // accepts in T CDATA: <span style=...> and <a href=...>. Everything
        // else is either mapped to a styled span (b/i/u/em/strong/code) or
        // unwrapped to its text.
        var sb = new StringBuilder();
        foreach (XmlNode child in node.ChildNodes)
        {
            if (child.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace)
            {
                sb.Append(WebUtility.HtmlEncode(child.Value));
                continue;
            }
            if (child is not XmlElement el) continue;
            var tag = el.LocalName.ToLowerInvariant();
            if (tag is "ul" or "ol") continue; // nested lists handled at block level
            var inner = ConvertInlineHtml(el);
            switch (tag)
            {
                case "b":
                case "strong": sb.Append($"<span style='font-weight:bold'>{inner}</span>"); break;
                case "i":
                case "em": sb.Append($"<span style='font-style:italic'>{inner}</span>"); break;
                case "u": sb.Append($"<span style='text-decoration:underline'>{inner}</span>"); break;
                case "code": sb.Append($"<span style='font-family:Consolas'>{inner}</span>"); break;
                case "br": sb.Append(' '); break;
                case "a":
                    var href = WebUtility.HtmlEncode(el.GetAttribute("href"));
                    if (href.Length > 0) sb.Append($"<a href=\"{href}\">{inner}</a>");
                    else sb.Append(inner);
                    break;
                case "span":
                    var style = el.GetAttribute("style").Replace('"', '\'');
                    if (style.Length > 0) sb.Append($"<span style='{style}'>{inner}</span>");
                    else sb.Append(inner);
                    break;
                default: sb.Append(inner); break;
            }
        }
        return sb.ToString();
    }

    private static XmlElement AddTextOE(XmlDocument doc, XmlNode parent, string inlineHtml)
    {
        var oe = doc.CreateElement("one", "OE", OneNS);
        var t = doc.CreateElement("one", "T", OneNS);
        t.AppendChild(doc.CreateCDataSection(ProtectCdata(inlineHtml)));
        oe.AppendChild(t);
        parent.AppendChild(oe);
        return oe;
    }

    private static void AddListItems(XmlDocument doc, XmlNode parent, XmlElement listNode, bool ordered)
    {
        foreach (XmlNode child in listNode.ChildNodes)
        {
            if (child is not XmlElement li || !li.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase)) continue;
            var oe = doc.CreateElement("one", "OE", OneNS);
            var list = doc.CreateElement("one", "List", OneNS);
            XmlElement marker;
            if (ordered)
            {
                marker = doc.CreateElement("one", "Number", OneNS);
                marker.SetAttribute("numberSequence", "0");
                marker.SetAttribute("numberFormat", "##.");
                marker.SetAttribute("language", "1033");
            }
            else
            {
                marker = doc.CreateElement("one", "Bullet", OneNS);
                marker.SetAttribute("bullet", "2");
                marker.SetAttribute("fontSize", "11.0");
            }
            list.AppendChild(marker);
            oe.AppendChild(list);
            var t = doc.CreateElement("one", "T", OneNS);
            t.AppendChild(doc.CreateCDataSection(ProtectCdata(ConvertInlineHtml(li))));
            oe.AppendChild(t);
            // A nested <ul>/<ol> inside the <li> becomes an OEChildren sub-list.
            XmlElement? nested = null;
            foreach (XmlNode sub in li.ChildNodes)
            {
                if (sub is XmlElement subEl && subEl.LocalName.ToLowerInvariant() is "ul" or "ol")
                {
                    if (nested is null)
                    {
                        nested = doc.CreateElement("one", "OEChildren", OneNS);
                        oe.AppendChild(nested);
                    }
                    AddListItems(doc, nested, subEl, subEl.LocalName.ToLowerInvariant() == "ol");
                }
            }
            parent.AppendChild(oe);
        }
    }

    private static void AddTableBlock(XmlDocument doc, XmlNode parent, XmlElement tableNode)
    {
        var rows = new List<XmlElement>();
        foreach (XmlNode row in tableNode.GetElementsByTagName("tr")) rows.Add((XmlElement)row);
        if (rows.Count == 0) return;
        var maxCells = 0;
        foreach (var row in rows)
        {
            var n = 0;
            foreach (XmlNode c in row.ChildNodes)
            {
                if (c is XmlElement cel && cel.LocalName.ToLowerInvariant() is "td" or "th") n++;
            }
            if (n > maxCells) maxCells = n;
        }
        var oe = doc.CreateElement("one", "OE", OneNS);
        var table = doc.CreateElement("one", "Table", OneNS);
        table.SetAttribute("bordersVisible", "true");
        var cols = doc.CreateElement("one", "Columns", OneNS);
        for (var i = 0; i < maxCells; i++)
        {
            var col = doc.CreateElement("one", "Column", OneNS);
            col.SetAttribute("index", i.ToString());
            col.SetAttribute("width", "160");
            cols.AppendChild(col);
        }
        table.AppendChild(cols);
        foreach (var row in rows)
        {
            var oneRow = doc.CreateElement("one", "Row", OneNS);
            var cellCount = 0;
            foreach (XmlNode cellNode in row.ChildNodes)
            {
                if (cellNode is not XmlElement cell) continue;
                var cellTag = cell.LocalName.ToLowerInvariant();
                if (cellTag is not ("td" or "th")) continue;
                var oneCell = doc.CreateElement("one", "Cell", OneNS);
                var cellKids = doc.CreateElement("one", "OEChildren", OneNS);
                var inline = ConvertInlineHtml(cell);
                if (cellTag == "th") inline = $"<span style='font-weight:bold'>{inline}</span>";
                AddTextOE(doc, cellKids, inline);
                oneCell.AppendChild(cellKids);
                oneRow.AppendChild(oneCell);
                cellCount++;
            }
            // Pad short rows so every row has the declared column count.
            while (cellCount < maxCells)
            {
                var oneCell = doc.CreateElement("one", "Cell", OneNS);
                var cellKids = doc.CreateElement("one", "OEChildren", OneNS);
                AddTextOE(doc, cellKids, "");
                oneCell.AppendChild(cellKids);
                oneRow.AppendChild(oneCell);
                cellCount++;
            }
            table.AppendChild(oneRow);
        }
        oe.AppendChild(table);
        parent.AppendChild(oe);
    }

    private static readonly Dictionary<int, string> HeadingSizes = new()
    {
        [1] = "20.0", [2] = "16.0", [3] = "13.0", [4] = "12.0", [5] = "11.0", [6] = "11.0",
    };

    private static void AddHtmlBlock(XmlDocument doc, XmlNode parent, XmlNode node)
    {
        if (node.NodeType is XmlNodeType.Text or XmlNodeType.CDATA)
        {
            var value = node.Value ?? "";
            if (value.Trim().Length > 0)
                AddTextOE(doc, parent, WebUtility.HtmlEncode(value.Trim()));
            return;
        }
        if (node is not XmlElement el) return;
        var tag = el.LocalName.ToLowerInvariant();
        var heading = Regex.Match(tag, "^h([1-6])$");
        if (heading.Success)
        {
            var size = HeadingSizes[int.Parse(heading.Groups[1].Value)];
            var inner = ConvertInlineHtml(el);
            AddTextOE(doc, parent, $"<span style='font-size:{size}pt;font-weight:bold'>{inner}</span>");
            return;
        }
        switch (tag)
        {
            case "p": AddTextOE(doc, parent, ConvertInlineHtml(el)); break;
            case "blockquote": AddTextOE(doc, parent, $"<span style='font-style:italic'>{ConvertInlineHtml(el)}</span>"); break;
            case "ul": AddListItems(doc, parent, el, false); break;
            case "ol": AddListItems(doc, parent, el, true); break;
            case "table": AddTableBlock(doc, parent, el); break;
            case "pre":
                foreach (var codeLine in Regex.Split(el.InnerText, "\r?\n"))
                    AddTextOE(doc, parent, $"<span style='font-family:Consolas'>{WebUtility.HtmlEncode(codeLine)}</span>");
                break;
            case "hr": AddTextOE(doc, parent, ""); break;
            default:
                // div/section/article and unknown containers: recurse block children.
                foreach (XmlNode child in el.ChildNodes) AddHtmlBlock(doc, parent, child);
                break;
        }
    }

    // Embedded images end up in notebooks that sync to OneDrive and may be
    // shared, so image_path must actually be an image — otherwise this tool is
    // an arbitrary-file-read that copies local files off the machine. Formats
    // are the ones OneNote's Image element accepts.
    private const long MaxImageBytes = 25 * 1024 * 1024;

    internal static bool LooksLikeSupportedImage(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 8 && (
            bytes[..8].SequenceEqual((ReadOnlySpan<byte>)new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }) // PNG
            || (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)                                               // JPEG
            || bytes[..4].SequenceEqual("GIF8"u8)                                                                       // GIF87a/89a
            || (bytes[0] == 0x42 && bytes[1] == 0x4D)                                                                   // BMP
            || bytes[..4].SequenceEqual((ReadOnlySpan<byte>)new byte[] { 0x49, 0x49, 0x2A, 0x00 })                      // TIFF LE
            || bytes[..4].SequenceEqual((ReadOnlySpan<byte>)new byte[] { 0x4D, 0x4D, 0x00, 0x2A }));                    // TIFF BE

    public Dictionary<string, object?> InsertRichContent(string pageId, string html, string imagePath) => Invoke(() =>
    {
        string xmlOut = "";
        App.GetPageContent(pageId, out xmlOut, PiBasic, Xs2013);
        var doc = Parse(xmlOut);
        var outline = doc.CreateElement("one", "Outline", OneNS);
        var children = doc.CreateElement("one", "OEChildren", OneNS);
        outline.AppendChild(children);
        if (html.Length > 0)
        {
            var prepared = PrepareXhtmlFragment(html);
            XmlDocument frag;
            try
            {
                frag = Parse($"<root>{prepared}</root>");
            }
            catch (XmlException ex)
            {
                throw new ArgumentException(
                    "html must be a well-formed XHTML fragment (self-close void tags, " +
                    $"match every open tag): {ex.Message}");
            }
            foreach (XmlNode child in frag.DocumentElement!.ChildNodes)
                AddHtmlBlock(doc, children, child);
            if (!children.HasChildNodes && imagePath.Length == 0)
                throw new ArgumentException("html fragment produced no content (empty or unsupported elements only)");
        }
        if (imagePath.Length > 0)
        {
            var info = new FileInfo(imagePath);
            if (!info.Exists) throw new ArgumentException($"image_path not found: {imagePath}");
            if (info.Length > MaxImageBytes)
                throw new ArgumentException(
                    $"image_path exceeds the {MaxImageBytes / (1024 * 1024)} MB embed limit: {imagePath}");
            var bytes = File.ReadAllBytes(imagePath);
            if (!LooksLikeSupportedImage(bytes))
                throw new ArgumentException(
                    $"image_path is not a supported image (png, jpeg, gif, bmp, or tiff): {imagePath}");
            var b64 = Convert.ToBase64String(bytes);
            var oe = doc.CreateElement("one", "OE", OneNS);
            var img = doc.CreateElement("one", "Image", OneNS);
            img.SetAttribute("format", "auto");
            var data = doc.CreateElement("one", "Data", OneNS);
            data.AppendChild(doc.CreateTextNode(b64));
            img.AppendChild(data);
            oe.AppendChild(img);
            children.AppendChild(oe);
        }
        if (!children.HasChildNodes) throw new ArgumentException("provide html and/or image_path");
        doc.DocumentElement!.AppendChild(outline);
        App.UpdatePageContent(doc.OuterXml, 0.0, Xs2013, false);
        return new Dictionary<string, object?> { ["page_id"] = pageId, ["inserted"] = true };
    });

    // --- Export --------------------------------------------------------------

    private static readonly Dictionary<string, string[]> ExportExtensions = new()
    {
        ["onenote"] = new[] { ".one" }, ["package"] = new[] { ".onepkg" },
        ["mhtml"] = new[] { ".mht", ".mhtml" }, ["pdf"] = new[] { ".pdf" },
        ["xps"] = new[] { ".xps" }, ["word"] = new[] { ".doc", ".docx" },
        ["docx"] = new[] { ".docx" }, ["emf"] = new[] { ".emf" },
        ["html"] = new[] { ".html", ".htm" },
    };

    // Exported content includes note text, which can carry injected
    // instructions from shared/clipped sources — so an unconstrained
    // path-and-extension write primitive could drop a script into an autorun
    // location. Pinning the extension to the declared format keeps the written
    // file inert, and requiring an existing absolute directory stops implicit
    // directory creation and relative-path surprises.
    internal static string ValidateExportTarget(string targetPath, string fmtName)
    {
        if (!Path.IsPathFullyQualified(targetPath))
            throw new ArgumentException($"target_path must be an absolute file path: {targetPath}");
        var full = Path.GetFullPath(targetPath);
        var dir = Path.GetDirectoryName(full);
        if (dir is null || !Directory.Exists(dir))
            throw new ArgumentException($"target_path directory does not exist: {dir}");
        var ext = Path.GetExtension(full).ToLowerInvariant();
        var allowed = ExportExtensions.GetValueOrDefault(fmtName, Array.Empty<string>());
        if (allowed.Length > 0 && !allowed.Contains(ext))
            throw new ArgumentException(
                $"target_path extension '{ext}' does not match format '{fmtName}' (expected {string.Join(" or ", allowed)})");
        return full;
    }

    public Dictionary<string, object?> Export(string objectId, string targetPath, string format) => Invoke(() =>
    {
        var fmtName = format.ToLowerInvariant();
        if (!PubFormat.TryGetValue(fmtName, out var fmt))
            throw new ArgumentException($"unknown export format: {fmtName} (pdf|html|docx|mhtml|xps|onenote)");
        var full = ValidateExportTarget(targetPath, fmtName);
        App.Publish(objectId, full, fmt, "");
        return new Dictionary<string, object?> { ["object_id"] = objectId, ["path"] = full, ["format"] = fmtName };
    });
}
