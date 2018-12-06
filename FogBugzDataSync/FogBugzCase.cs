using System;
using System.Collections.Generic;
using System.Text;

namespace Inflectra.SpiraTest.PlugIns.FogBugzDataSync
{
    public class FogBugzCase
    {
        protected string title = "";
        protected int project = -1;
        protected int area = -1;
        protected int category = -1;
        protected string computer = "";
        protected string description = "";
        protected DateTime? due;
        protected DateTime? closed;
        protected DateTime? lastUpdated;
        protected int fixFor = -1;
        protected int hrsCurrEst = -1;
        protected int id = -1;
        protected int personAssignedTo = -1;
        protected int personOpenedBy = -1;
        protected int priority = -1;
        protected int status = -1;
        protected string version = "";

        public string Title
        {
            get
            {
                return this.title;
            }
            set
            {
                this.title = value;
            }
        }

        public int Project
        {
            get
            {
                return this.project;
            }
            set
            {
                this.project = value;
            }
        }

        public int Area
        {
            get
            {
                return this.area;
            }
            set
            {
                this.area = value;
            }
        }

        public int FixFor
        {
            get
            {
                return this.fixFor;
            }
            set
            {
                this.fixFor = value;
            }
        }

        public int Category
        {
            get
            {
                return this.category;
            }
            set
            {
                this.category = value;
            }
        }

        /// <summary>
        /// ixPersonOpenedBy
        /// </summary>
        public int PersonOpenedBy
        {
            get
            {
                return this.personOpenedBy;
            }
            set
            {
                this.personOpenedBy = value;
            }
        }

        public int PersonAssignedTo
        {
            get
            {
                return this.personAssignedTo;
            }
            set
            {
                this.personAssignedTo = value;
            }
        }

        public int Priority
        {
            get
            {
                return this.priority;
            }
            set
            {
                this.priority = value;
            }
        }

        public int Status
        {
            get
            {
                return this.status;
            }
            set
            {
                this.status = value;
            }
        }

        public DateTime? Due
        {
            get
            {
                return this.due;
            }
            set
            {
                this.due = value;
            }
        }

        public DateTime? Closed
        {
            get
            {
                return this.closed;
            }
            set
            {
                this.closed = value;
            }
        }

        public DateTime? LastUpdated
        {
            get
            {
                return this.lastUpdated;
            }
            set
            {
                this.lastUpdated = value;
            }
        }

        public int HrsCurrEst
        {
            get
            {
                return this.hrsCurrEst;
            }
            set
            {
                this.hrsCurrEst = value;
            }
        }

        public string Version
        {
            get
            {
                return this.version;
            }
            set
            {
                this.version = value;
            }
        }

        public string Computer
        {
            get
            {
                return this.computer;
            }
            set
            {
                this.computer = value;
            }
        }

        public string Description
        {
            get
            {
                return this.description;
            }
            set
            {
                this.description = value;
            }
        }

        public int Id
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }
    }
}
