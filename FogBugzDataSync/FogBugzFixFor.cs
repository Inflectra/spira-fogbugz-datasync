using System;
using System.Collections.Generic;
using System.Text;

namespace Inflectra.SpiraTest.PlugIns.FogBugzDataSync
{
    /// <summary>
    /// Represents a single fix-for in FogBugz
    /// </summary>
    public class FogBugzFixFor
    {
        /// <summary>
        /// The id of the fix-for
        /// </summary>
        public int Id
        {
            get;
            set;
        }

        public int Project
        {
            get;
            set;
        }

        public string Name
        {
            get;
            set;
        }

        public DateTime ReleaseDate
        {
            get;
            set;
        }

        public bool Assignable
        {
            get;
            set;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public FogBugzFixFor()
        {
            this.Id = -1;
            this.Project = -1;
            this.Name = "";
            this.Assignable = true;
        }
    }
}
