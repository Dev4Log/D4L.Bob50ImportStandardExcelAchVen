using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dev4LogLayer.Bob50.Models;

namespace D4L.Bob50ImportStandardExcelAchVen.Models
{
    public class DmInvDoc
    {
        string fyear;
        int year;
        int month;
        string dbk;
        string dbtype;
        int numdoc;
        string modifiedBy;
        string createdBy;
        string modifiedOn;
        string createdOn;
        Bob50Folder folder;

        public DmInvDoc()
        {
        }

        public DmInvDoc(string fyear, int year, int month, string dbk, string dbtype, int numdoc, string modifiedBy, string createdBy, string modifiedOn, string createdOn, Bob50Folder folder)
        {
            this.fyear = fyear;
            this.year = year;
            this.month = month;
            this.dbk = dbk;
            this.dbtype = dbtype;
            this.numdoc = numdoc;
            this.modifiedBy = modifiedBy;
            this.createdBy = createdBy;
            this.modifiedOn = modifiedOn;
            this.createdOn = createdOn;
            this.folder = folder;
        }

        public string Fyear { get => fyear; set => fyear = value; }
        public int Year { get => year; set => year = value; }
        public int Month { get => month; set => month = value; }
        public string Dbk { get => dbk; set => dbk = value; }
        public string Dbtype { get => dbtype; set => dbtype = value; }
        public int Numdoc { get => numdoc; set => numdoc = value; }
        public string ModifiedBy { get => modifiedBy; set => modifiedBy = value; }
        public string CreatedBy { get => createdBy; set => createdBy = value; }
        public string ModifiedOn { get => modifiedOn; set => modifiedOn = value; }
        public string CreatedOn { get => createdOn; set => createdOn = value; }
        public Bob50Folder Folder { get => folder; set => folder = value; }
    }
}
