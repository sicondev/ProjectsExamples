using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Coordinators;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectTransactions.Coordinators;
using Sicon.Sage200.Projects.Objects.Instruments;
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

        /// <summary>
        /// Create Project transaction for POP invoice and reconcile
        /// </summary>
        /// <param name="oPOPOrderReturnLine"></param>
        /// <param name="oPLEntry"></param>
        /// <param name="oProject"></param>
        /// <param name="oProjectHeader"></param>
        /// <param name="PhaseID"></param>
        /// <param name="StageID"></param>
        /// <param name="ActivityID"></param>
        /// <param name="oNLCode"></param>
        /// <param name="ThisNarrative"></param>
        /// <param name="LineTotalValue"></param>
        public void ReconcilePOPInvoiceToProjectAndNominal(Sage.Accounting.POP.POPOrderReturnLine oPOPOrderReturnLine, Sage.Accounting.PurchaseLedger.PostedPurchaseAccountEntry oPLEntry, SiJcJob oProject, SiJcChd oProjectHeader, long PhaseID, long StageID, long ActivityID,
            Sage.Accounting.NominalLedger.NominalCode oNLCode, string ThisNarrative, decimal LineTotalValue)
        {
            try
            {
                string TranType = "POPINV";
                if (oPLEntry.EntryType == Sage.Accounting.TradingAccountEntryTypeEnum.TradingAccountEntryTypeCreditNote)
                {
                    TranType = "POPCredit";
                }
                ProjectAnalysisItem oProjectAnalysisItem = ProjectAnalysisItemFactory.Factory.CreateNew();

                //Populate fields for POP document type transaction
                oProjectAnalysisItem.Generate(oProject, oProjectHeader, oPLEntry.InstrumentDate, oNLCode, ThisNarrative, LineTotalValue, "", TranType, oPLEntry.URN, oPLEntry.PLPostedSupplierTran, true, false, false, PhaseID, StageID, ActivityID);

                //Fill in some extra fields
                oProjectAnalysisItem.ThisQuantity = 1;
                oProjectAnalysisItem.Quantity = 1;

                oProjectAnalysisItem.Supplier = oPLEntry.Supplier;
                oProjectAnalysisItem.POPOrderReturnLine = oPOPOrderReturnLine;
                oProjectAnalysisItem.POPOrderReturn = oPOPOrderReturnLine.POPOrderReturn;
                oProjectAnalysisItem.PLEntry = oPLEntry;

                //Post transaction creating SIJCTRN and reconciling to nominal
                oProjectAnalysisItem.Post();

                //Update fields after post (Optional)
                oProjectAnalysisItem.SiJcTrn.InvoiceCreditNumber = oPLEntry.InstrumentNo;
                oProjectAnalysisItem.SiJcTrn.InvoiceCreditDate = oPLEntry.InstrumentDate;
                oProjectAnalysisItem.SiJcTrn.Update();

            }
            catch (Exception)
            {

                throw;
            }
        }


        public void PostPOPInvoiceAndCreateProjectTransaction()
        {
            try
            {
                PostPOPInvoiceCoordinator oCoordinator = new PostPOPInvoiceCoordinator()
                {
                    SupplierReference = "ABC001",//Supplier Reference
                    InvoiceCreditDate = DateTime.Today,//Invoice date
                    InvoiceCreditNo = "INV008",//Invoice number
                    SecondReference = "2nd Ref",//Second Reference on Invoices
                    //OPTIONAL: add invoice value (*Variance will be added as an extra nominal line)
                    InvCredGoodsValue = 120,
                    InvCredTaxValue = 25,
                };

                //Loop Through Orders and lines available to invoice

                POPInvoiceItem oInstrument = null;
                oInstrument = new POPInvoiceItem()
                {
                    OrderDocumentNo = "0000000066"//POP Order Number
                };

                
                oInstrument.POPLineIDs.Add(Convert.ToInt64(515123)); //POP order line ID
                oCoordinator.POPInvoiceItems.Add(oInstrument);


                //Post
                long URN = oCoordinator.PostInvoice();

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
