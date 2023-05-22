using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Coordinators;
using Sicon.Sage200.Projects.Objects.Factory;
using System;

namespace ProjectsExamples
{
    public class POPMethods
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
        /// Get a project transaction (SIJCTRN) for a POP Order Return Line
        /// </summary>
        /// <param name="POPOrderReturnLineID"></param>
        public void GetProjectTransactionForPOP(long POPOrderReturnLineID)
        {
            try
            {
                //Get project transaction by <long> POPOrderReturnLineID
                SiJcTrn oPOP_SIJCTRN = ProjectTransactionFactory.Factory.FetchPOP(POPOrderReturnLineID);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Simple method to update a POP transaction in projects
        /// </summary>
        /// <param name="oPOPOrderReturnLine"></param>
        /// <param name="oProject">See Methods to retrieve project</param>
        /// <param name="oProjectHeader">See Methods to retrieve Project Header</param>
        /// <param name="PhaseID">See Methods to retrieve Project Levels</param>
        /// <param name="StageID">See Methods to retrieve Project Levels</param>
        /// <param name="ActivityID">See Methods to retrieve Project Levels</param>
        public void CreateOrAmendProjectTransactionForPOP(Sage.Accounting.POP.POPOrderReturnLine oPOPOrderReturnLine, SiJcJob oProject, SiJcChd oProjectHeader, long PhaseID, long StageID, long ActivityID)
        {
            try
            {
                //Update POP SIJCTRN transaction
                POPPostingCoordinator.UpdatePOPLineTransaction(oPOPOrderReturnLine, oProject, oProjectHeader, PhaseID, StageID, ActivityID, 0, 0);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Example of Setting Analysis on Order Header
        /// </summary>
        /// <param name="POPOrderReturnID"></param>
        /// <param name="ProjectNumber"></param>
        /// <param name="ProjectHeaderID">See Methods to retrieve Project Header</param>
        /// <param name="PhaseID">See Methods to retrieve Project Levels</param>
        /// <param name="StageID">See Methods to retrieve Project Levels</param>
        /// <param name="ActivityID">See Methods to retrieve Project Levels</param>
        public void UpdateProjectAnalysisOnOrderHeader(long POPOrderReturnID, string ProjectNumber, long ProjectHeaderID, long PhaseID, long StageID, long ActivityID)
        {
            try
            {
                //Update project analysis on the order header
                POPPostingCoordinator.UpdateOrderHeaderProjectAnalysis(POPOrderReturnID, ProjectNumber, ProjectHeaderID, PhaseID, StageID, ActivityID, 0);
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
