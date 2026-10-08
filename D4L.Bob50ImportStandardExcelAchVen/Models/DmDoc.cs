using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dev4LogLayer.Bob50.Models;

namespace D4L.Bob50ImportStandardExcelAchVen.Models
{
    public class DmDoc
    {
        string fileName;
        Bob50Folder folder;
        string createdBy;
        string createdOn;
        string modifiedBy;
        string modifiedOn;
        string cid;
        DateTime dateDoc;
        DmInvDoc dmInvDoc;

        public DmDoc()
        {
        }

        public DmDoc(string fileName, Bob50Folder folder, string createdBy, string createdOn, string modifiedBy, string modifiedOn, string cid, DateTime dateDoc, DmInvDoc dmInvDoc)
        {
            this.FileName = fileName;
            this.Folder = folder;
            this.CreatedBy = createdBy;
            this.CreatedOn = createdOn;
            this.ModifiedBy = modifiedBy;
            this.ModifiedOn = modifiedOn;
            this.Cid = cid;
            this.DateDoc = dateDoc;
            this.DmInvDoc = dmInvDoc;
        }

        public string FileName { get => fileName; set => fileName = value; }
        public Bob50Folder Folder { get => folder; set => folder = value; }
        public string CreatedBy { get => createdBy; set => createdBy = value; }
        public string CreatedOn { get => createdOn; set => createdOn = value; }
        public string ModifiedBy { get => modifiedBy; set => modifiedBy = value; }
        public string ModifiedOn { get => modifiedOn; set => modifiedOn = value; }
        public string Cid { get => cid; set => cid = value; }
        public DateTime DateDoc { get => dateDoc; set => dateDoc = value; }
        public DmInvDoc DmInvDoc { get => dmInvDoc; set => dmInvDoc = value; }
    }
}
