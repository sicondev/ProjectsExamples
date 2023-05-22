using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectHeaders.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.Projects.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectTransactions.Coordinators;
using System;

namespace ProjectsExamples
{
    public class ProjectHeaderMethods
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
        /// Get a project header by ID
        /// </summary>
        /// <param name="ProjectHeaderID"></param>
        public void GetProjectHeader(long ProjectHeaderID)
        {
            try
            {
                //Get project header by header code
                SiJcChd oProjectHeader = ProjectHeaderFactory.Factory.Fetch(ProjectHeaderID);

            }
            catch (Exception)
            {

                throw;
            }
        }
        /// <summary>
        /// Get a project header by Code
        /// </summary>
        /// <param name="ProjectHeaderCode">For Example 'Materials'</param>
        public void GetProjectHeader(string ProjectHeaderCode)
        {
            try
            {
                //Get project header by header code
                SiJcChd oProjectHeader = ProjectHeaderFactory.Factory.FetchWithCode(ProjectHeaderCode);

                //*Note Sales orders should only use project headers with the ‘HeaderType’ = ‘Revenue’ and Purchase orders use 'Cost' codes

            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
