using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteStats
{
    internal sealed class NotebookStats
    {
        private const string OneNoteXmlNamespace = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

        private OneNote.Application _onenoteApp;
        private XmlDocument _notebookXmlDoc;
        private XmlNamespaceManager _nsManager;

        public string NotebookName { get; private set; }
        public int SectionGroupCount { get { return getSectionGroupCount(); } }
        public int SectionCount { get { return getSectionCount(); } }
        public int PageCount { get { return getPageCount(); } }

        public NotebookStats(string notebookName)
        {
            NotebookName = notebookName;
            _onenoteApp = new OneNote.Application();

            initializeNotebookXmlDoc();
        }

        private void initializeNotebookXmlDoc()
        {
            string notebookNodeId = getNotebookNodeId(_onenoteApp, NotebookName);

            string hierarchyXml;
            _onenoteApp.GetHierarchy(notebookNodeId, OneNote.HierarchyScope.hsPages, out hierarchyXml);

            _notebookXmlDoc = new XmlDocument();
            _notebookXmlDoc.LoadXml(hierarchyXml);

            _nsManager = new XmlNamespaceManager(_notebookXmlDoc.NameTable);
            _nsManager.AddNamespace(@"one", OneNoteXmlNamespace);
        }

        public List<PageInfo> GetPageInfo()
        {
            List<PageInfo> pageInfos = new List<PageInfo>();

            foreach (XmlNode pageNode in getAllPageNodes())
            {
                PageInfo pageInfo = new PageInfo()
                {
                    Id = pageNode.Attributes[@"ID"].Value,
                    Name = pageNode.Attributes[@"name"].Value,
                    DateTime = DateTime.Parse(pageNode.Attributes[@"dateTime"].Value),
                    LastModifiedTime = DateTime.Parse(pageNode.Attributes[@"lastModifiedTime"].Value),
                    PageLevel = Int32.Parse(pageNode.Attributes[@"pageLevel"].Value),
                };

                if (pageNode.Attributes[@"isCurrentlyViewed"] != null)
                {
                    pageInfo.IsCurrentlyViewed = pageNode.Attributes[@"isCurrentlyViewed"].Value;
                }

                pageInfo.LocationPath = getLocationPath(pageNode);

                pageInfos.Add(pageInfo);
            }

            return pageInfos;
        }

        private string getLocationPath(XmlNode node)
        {
            if (node == null) throw new ArgumentNullException(@"node");
            if ((node.NodeType != XmlNodeType.Element) && (node.NodeType != XmlNodeType.Document)) throw new ArgumentOutOfRangeException(@"node", node.NodeType, @"The node is not element node or document node.");

            string parentPath = string.Empty;
            if (node.ParentNode != null)
            {
                parentPath = getLocationPath(node.ParentNode);
            }

            if ((node.Attributes != null) && (node.Attributes[@"name"].Value != null))
            {
                string name = node.Attributes[@"name"].Value;
                return parentPath + Path.DirectorySeparatorChar + name;
            }

            return parentPath;
        }

        private int getSectionGroupCount()
        {
            return getAllSectionGroupNodes().Count;
        }

        private XmlNodeList getAllSectionGroupNodes()
        {
            return _notebookXmlDoc.SelectNodes(@"(//one:SectionGroup[@name!='OneNote_RecycleBin']|one:SectionGroup[@name!='OneNote_RecycleBin']//one:SectionGroup)", _nsManager);
        }

        private int getSectionCount()
        {
            return getAllSectionNodes().Count;
        }

        private XmlNodeList getAllSectionNodes()
        {
            return _notebookXmlDoc.SelectNodes(@"(//one:Section|one:SectionGroup[@name!='OneNote_RecycleBin']//one:Section)", _nsManager);
        }

        private int getPageCount()
        {
            return getAllPageNodes().Count;
        }

        private XmlNodeList getAllPageNodes()
        {
            return _notebookXmlDoc.SelectNodes(@"(//*/one:Page|one:SectionGroup[@name!='OneNote_RecycleBin']//one:Page)", _nsManager);
        }

        private static string getNotebookNodeId(OneNote.Application onenoteApp, string notebookName)
        {
            string notebooksXml;
            onenoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsNotebooks, out notebooksXml);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(notebooksXml);

            XmlNamespaceManager nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
            nsManager.AddNamespace(@"one", OneNoteXmlNamespace);

            string notebookXpath = getNotebookXpath(notebookName);
            XmlNode notebookNode = xmlDoc.SelectSingleNode(notebookXpath, nsManager);

            if (notebookNode == null)
            {
                throw new ArgumentException(string.Format(@"Can not found ""{0}"" as notebook name.", notebookName), @"notebookName");
            }

            return notebookNode.Attributes[@"ID"].Value;
        }

        private static string getNotebookXpath(string notebookName)
        {
            return string.Format(@"//one:Notebook[@nickname='{0}']", notebookName);
        }
    }
}
