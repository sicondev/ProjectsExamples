using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Infrastructure.Projects.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectTransactions.Coordinators;
using System;

namespace ProjectsExamples
{
    public class SOPMethods
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
        /// Get a project transaction (SIJCTRN) for a SOP Order Return Line
        /// </summary>
        /// <param name="SOPOrderReturnLineID"></param>
        public void GetProjectTransactionForSOP(long SOPOrderReturnLineID)
        {
            try
            {
                //Get project transaction by <long> SOPOrderReturnLineID
                SiJcTrn oSOP_SIJCTRN = ProjectTransactionFactory.Factory.FetchSOP(SOPOrderReturnLineID);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Simple method to update a SOP transaction in projects
        /// </summary>
        /// <param name="oSOPOrderReturnLine"></param>
        /// <param name="oProject">See Methods to retrieve project</param>
        /// <param name="oProjectHeader">See Methods to retrieve Project Header</param>
        /// <param name="PhaseID">See Methods to retrieve Project Levels</param>
        /// <param name="StageID">See Methods to retrieve Project Levels</param>
        /// <param name="ActivityID">See Methods to retrieve Project Levels</param>
        public void CreateOrAmendProjectTransactionForSOP(Sage.Accounting.SOP.SOPOrderReturnLine oSOPOrderReturnLine, SiJcJob oProject, SiJcChd oProjectHeader, long PhaseID, long StageID, long ActivityID)
        {
            try
            {
                //Update SOP SIJCTRN transaction
                SOPPostingCoordinator.UpdateSOPLineTransaction(oSOPOrderReturnLine, oProject, oProjectHeader, PhaseID, StageID, ActivityID, 0);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Example of Setting Analysis on Order Header
        /// </summary>
        /// <param name="SOPOrderReturnID"></param>
        /// <param name="ProjectNumber"></param>
        /// <param name="ProjectHeaderID">See Methods to retrieve Project Header</param>
        /// <param name="PhaseID">See Methods to retrieve Project Levels</param>
        /// <param name="StageID">See Methods to retrieve Project Levels</param>
        /// <param name="ActivityID">See Methods to retrieve Project Levels</param>
        public void UpdateProjectAnalysisOnOrderHeader(long SOPOrderReturnID, string ProjectNumber, long ProjectHeaderID, long PhaseID, long StageID, long ActivityID)
        {
            try
            {
                //Update project analysis on the order header
                SOPPostingCoordinator.UpdateOrderHeaderProjectAnalysis(SOPOrderReturnID, ProjectNumber, ProjectHeaderID, PhaseID, StageID, ActivityID);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Method for creating and linking up project transactions for SOP invoices. (Requires v221.0.18+)
        /// </summary>
        /// <param name="oInvoice">Sage.Accounting.SOP.SOPInvoiceCredit object</param>
        /// <param name="PostedEntryID">Sage.Accounting.SalesLedger.PostedSalesAccountEntry ID</param>
        public void PostProjectTransactionsForSOPInvoices(Sage.Accounting.SOP.SOPInvoiceCredit oInvoice, long PostedEntryID)
        {
            try
            {
                //Post Invoice method to create project transactions and link to nominals
                SOPPostingCoordinator.PostSOPInvoices(oInvoice, PostedEntryID);
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
