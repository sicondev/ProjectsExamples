using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectHeaders.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.Projects.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectTransactions.Coordinators;
using System;

namespace ProjectsExamples
{
    public class ProjectLevelMethods
    {
        ////Code Examples valid in latest hotfixes of supported versions (v21.1+)
        ///////////////////////////////////////////////////////////////////////////////////////////////////
        ////The Projects Library can be found within the assembly cache when running sage once installed.
        ////C:\Users\<User Name>\AppData\Local\Sage\Sage200\AssemblyCache
        ////The.dlls needed are:
        ////Sicon.Sage200.Projects.Objects
        ////Sicon.API.Sage200.Objects
        ///////////////////////////////////////////////////////////////////////////////////////////////////


        /// <summary>
        /// Method to retrieve Project Level
        /// </summary>
        /// <param name="oProject">See methods for retrieving projects</param>
        /// <param name="PhaseDescription">For example 'Phase1'</param>
        /// <param name="StageDescription">For example 'Stage1'</param>
        /// <param name="ActivityDescription">For example 'Activity1'</param>
        public void GetProjectLevel(SiJcJob oProject, string PhaseDescription, string StageDescription, string ActivityDescription)
        {
            try
            {
                //Get Project levels by codes (Will return default if not found)
                ProjectLevelFactory.ProjectLevel oProjectLevel = ProjectLevelFactory.GetProjectLevel(oProject, PhaseDescription, StageDescription, ActivityDescription);
                long PhaseID = oProjectLevel.PhaseID;
                long StageID = oProjectLevel.StageID;
                long ActivityID = oProjectLevel.ActivityID;

            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
