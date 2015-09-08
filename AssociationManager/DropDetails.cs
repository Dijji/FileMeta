using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using GongSolutions.Wpf.DragDrop;

namespace FileMetadataAssociationManager
{
    class DropDetails
    {
        private bool isTags = false;
        private bool isFiles = false;
        private bool hasParentTag = false;
        private IEnumerable<TreeItem> items = null;

        public bool IsEmpty { get { return !isTags && !isFiles; } }
        public bool IsTags { get { return isTags; } }
        public bool HasParentTag { get { return hasParentTag; } }
        public bool IsFiles { get { return isFiles; } }

        public IEnumerable<TreeItem> Items { get { return items; } }

        public Controls Source { get; set; }
        public Controls Target { get; set; }

        //public IEnumerable<FileNode> Files
        //{
        //    get
        //    {
        //        return items.Select(ti => (FileNode)ti.Node);
        //    }
        //}

        //public IEnumerable<TagNode> Tags
        //{
        //    get
        //    {
        //        return items.Select(ti => (TagNode)ti.Node);
        //    }
        //}

        //public string Description
        //{
        //    get
        //    {
        //        return View.NameList(IsFiles ? "file" : "tag", items);
        //    }
        //}

        //public string ChildDescription
        //{
        //    get
        //    {
        //        return items.Count() > 1 ? "as children" : "as a child";
        //    }
        //}

        //public bool AnyFileHasTag(TagNode tag)
        //{
        //    foreach (FileNode f in Files)
        //        if (f.Tags.Contains(tag))
        //            return true;

        //    return false;
        //}

        //public bool AllFilesHaveTag(TagNode tag)
        //{
        //    foreach (FileNode f in Files)
        //        if (!f.Tags.Contains(tag))
        //            return false;

        //    return true;
        //}


        public DropDetails(IDropInfo dropInfo)
        {


            //if (dropInfo.Data is TreeItem)
            //{
            //    TreeItem from = (TreeItem)dropInfo.Data;
            //    items = new TreeItem[] { from };
            //    if (from.Node is TagNode)
            //    {
            //        isTags = true;
            //        hasParentTag = ((TagNode)from.Node).Parents.Count > 0;
            //    }
            //    else
            //        isFiles = from.Node is FileNode;
            //}
            //else if (dropInfo.Data is IEnumerable<TreeItem>)
            //{
            //    bool inconsistent = ((IEnumerable<TreeItem>)dropInfo.Data).DistinctBy(ti => ti.Node.GetType()).Count() > 1;

            //    if (!inconsistent)
            //    {
            //        items = (IEnumerable<TreeItem>)dropInfo.Data;
            //        if (items.First().Node is TagNode)
            //        {
            //            isTags = true;
            //            hasParentTag = ((TagNode)items.First().Node).Parents.Count > 0;
            //        }
            //        else
            //            isFiles = items.First().Node is FileNode;
            //    }
            //}
        }



    }
}
