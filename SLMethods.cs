using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Infrastructure.ProjectHeaders.Factories;
using Sicon.Sage200.Projects.Objects.Infrastructure.Projects.Factories;
using Sicon.Sage200.Projects.Objects.Instruments;
using System;
using System.Collections.Generic;

namespace ProjectsExamples
{
    public class SLMethods
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
        /// Get a project transactions (SIJCTRN) for a PL Invoice Line
        /// </summary>
        /// <param name="SLPostedTransactionID"></param>
        public void GetProjectTransactionsForSLInvoice(long SLPostedTransactionID)
        {
            try
            {
                //Get project transactions by <long> PLPostedTransactionID
                SiJcTrns oSL_SIJCTRNs = ProjectTransactionsFactory.Factory.FetchSL(SLPostedTransactionID, "SLINV", true);
            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// Post And Reconcile SL Invoice To Project
        /// </summary>
        /// <param name="oCustomer"></param>
        /// <param name="DueDate"></param>
        /// <param name="PostedDate"></param>
        /// <param name="TransactionDate"></param>
        /// <param name="QueryCode"></param>
        /// <param name="DaysDiscountValid"></param>
        /// <param name="DiscountPercentage"></param>
        /// <param name="DiscountValue"></param>
        /// <param name="NominalReference"></param>
        /// <param name="NominalCostCentre"></param>
        /// <param name="NominalDepartment"></param>
        /// <param name="NominalNarrative"></param>
        /// <param name="NominalAmount"></param>
        /// <param name="NominalAnalysisCode"></param>
        /// <param name="TaxRateNumber"></param>
        /// <param name="TaxAmountOnLine"></param>
        /// <param name="ProjectNumber"></param>
        /// <param name="ProjectHeaderCode"></param>
        /// <param name="PhaseDescription"></param>
        /// <param name="StageDescription"></param>
        /// <param name="ActivityDescription"></param>
        public void PostAndReconcileSLInvoiceToProject(Sage.Accounting.SalesLedger.Customer oCustomer, DateTime DueDate, DateTime PostedDate, DateTime TransactionDate, string QueryCode, short DaysDiscountValid, decimal DiscountPercentage, decimal DiscountValue,
            string NominalReference, string NominalCostCentre, string NominalDepartment, string NominalNarrative, decimal NominalAmount, string NominalAnalysisCode, short TaxRateNumber, decimal TaxAmountOnLine, 
            string ProjectNumber, string ProjectHeaderCode, string PhaseDescription, string StageDescription, string ActivityDescription)
        {
            List<ProjectAnalysisItem> _SiJCAnalysisItems = null;
            SiJcChd oSiJcChd = null;
            SiJcJob oSiJcJob = null;
            Sage.Accounting.SalesLedger.SalesInvoiceInstrument oSalesInvoiceInstrument = null;
            try
            {
                //Create New Invoice
                Sicon.API.Sage200.Objects.SalesLedger.SalesInvoiceDetails oSalesInvoiceDetails = new Sicon.API.Sage200.Objects.SalesLedger.SalesInvoiceDetails(oCustomer.Reference);
                oSalesInvoiceInstrument = Sicon.API.Sage200.Objects.StaticClass.CreateSalesInvoiceInstrument();
                oSalesInvoiceInstrument.Customer = oCustomer;
                oSalesInvoiceInstrument.NominalAnalysisItems.Empty();
                oSalesInvoiceInstrument.TaxAnalysisItems.Empty();

                oSalesInvoiceInstrument.SuppressExceedsCreditLimitException = true;// oSalesInvoiceDetails.SuppressExceedsCreditLimitException;
                oSalesInvoiceInstrument.SuppressInvalidDueDateException = true;

                //Set the Invoice Numbers
                oSalesInvoiceInstrument.InstrumentNo = "Ref1";
                oSalesInvoiceInstrument.SecondReferenceNo = "Ref2";

                //set the dates
                oSalesInvoiceInstrument.DueDate = DueDate;
                oSalesInvoiceInstrument.PostedDate = PostedDate;
                oSalesInvoiceInstrument.InstrumentDate = TransactionDate;
                if (oSalesInvoiceInstrument.DueDate < oSalesInvoiceInstrument.InstrumentDate)
                {
                    oSalesInvoiceInstrument.DueDate = oSalesInvoiceInstrument.InstrumentDate;
                }

                // set query code
                oSalesInvoiceInstrument.Queried = QueryCode;

                //set the discount
                oSalesInvoiceInstrument.DiscountDays = DaysDiscountValid;
                oSalesInvoiceInstrument.DiscountPercent = DiscountPercentage;
                oSalesInvoiceInstrument.DiscountValue = DiscountValue;

                //Loop through nominals if multiple example is for one nominal. 
                //NOTE!! Each nominal will be per One Project transaction
                Sage.Accounting.Common.NominalSpecification oNLSpec = Sage.Accounting.Common.NominalSpecificationFactory.Factory.CreateNew(NominalReference, NominalCostCentre, NominalDepartment);
                Sage.Accounting.NominalLedger.NominalCode oNLCode = Sage.Accounting.NominalLedger.NominalCodeFactory.Factory.Fetch(oNLSpec);

                Sage.Accounting.TradeLedger.NominalAnalysisItem oNominalAnalysisItem = (Sage.Accounting.TradeLedger.NominalAnalysisItem)oSalesInvoiceInstrument.NominalAnalysisItems.AddNew();
                oNominalAnalysisItem.SetNominalCodeAndSpecification(oNLCode, oNLCode.NominalSpecification);
                oNominalAnalysisItem.Narrative = NominalNarrative;
                oNominalAnalysisItem.Amount = Math.Round(NominalAmount, 2);
                oNominalAnalysisItem.TransactionAnalysisCode = NominalAnalysisCode;

                //Invoice Tax line
                Sage.Accounting.TradeLedger.TaxAnalysisItem oTaxAnalysisItem = null;
                foreach (Sage.Accounting.TradeLedger.TaxAnalysisItem ThisTaxAnalysisItem in oSalesInvoiceInstrument.TaxAnalysisItems)
                {
                    if (ThisTaxAnalysisItem.TaxCode.Code == TaxRateNumber)
                    {
                        oTaxAnalysisItem = ThisTaxAnalysisItem;
                        break;
                    }
                }
                if (oTaxAnalysisItem == null)
                {
                    oTaxAnalysisItem = (Sage.Accounting.TradeLedger.TaxAnalysisItem)oSalesInvoiceInstrument.TaxAnalysisItems.AddNew();
                }

                Sage.Accounting.TaxModule.TaxCodes oTaxCodes = Sage.Accounting.TaxModule.TaxCodesFactory.Factory.CreateNew();
                oTaxCodes.Query.Filters.Add(new Sage.ObjectStore.Filter(Sage.Accounting.TaxModule.TaxCode.FIELD_CODE, TaxRateNumber));
                oTaxCodes.Find();
                if (oTaxCodes.First != null)
                {
                    oTaxAnalysisItem.TaxCode = oTaxCodes.First;
                }
                else
                {
                    oTaxAnalysisItem.TaxCode = Sage.Accounting.TaxModule.TaxCodeFactory.Factory.FetchFirstNonZeroRatedTaxCode();
                }

                //Reset nominal value as it may have been changed when setting the tax code
                oNominalAnalysisItem.Amount = Math.Round(NominalAmount, 2);

                decimal TaxAmount = oTaxAnalysisItem.TaxAmount;
                oTaxAnalysisItem.Goods += Math.Round(NominalAmount, 2);
                oTaxAnalysisItem.TaxAmount = TaxAmount + Math.Round(TaxAmountOnLine, 2);

                //Project Transaction Line
                oSiJcJob = ProjectFactory.Factory.FetchWithProjectNumber(ProjectNumber, true);
                oSiJcChd = ProjectHeaderFactory.Factory.FetchWithCode(ProjectHeaderCode, true);


                #region Phases/Stages/Activites
                ProjectHierachyInstrument oJHInstrument = new ProjectHierachyInstrument();
                oJHInstrument.Load(oSiJcJob.SiJcJob, PhaseDescription, StageDescription, ActivityDescription);
                long PhaseID = oJHInstrument.PhaseID;
                long StageID = oJHInstrument.StageID;
                long ActivityID = oJHInstrument.ActivityID;
                #endregion

                //"SLINV" = Invoice
                //"SLCREDIT" = Credit Note

                ProjectAnalysisItem oAnalysisItem = ProjectAnalysisItemFactory.Factory.CreateNew();
                oAnalysisItem.Generate(oSalesInvoiceInstrument, oSiJcJob, oSiJcChd, oNominalAnalysisItem.NominalCode, oNominalAnalysisItem.Narrative, oNominalAnalysisItem.Amount, oNominalAnalysisItem.TransactionAnalysisCode, "SLINV", oNominalAnalysisItem, true, false, false, PhaseID, StageID, ActivityID, 0, 0);
                oAnalysisItem.TranDate = oSalesInvoiceInstrument.InstrumentDate;
                _SiJCAnalysisItems.Add(oAnalysisItem);


                //Final Posting after Lines are populated
                decimal TaxTotal = oSalesInvoiceInstrument.TaxAmountTotal;
                oSalesInvoiceInstrument.NetValue = oSalesInvoiceInstrument.NominalAnalysisTotal;
                oSalesInvoiceInstrument.TaxValue = TaxTotal;

                oSalesInvoiceInstrument.Validate();
                oSalesInvoiceInstrument.Update();

                foreach (ProjectAnalysisItem pendingItem in _SiJCAnalysisItems)
                {
                    if (pendingItem.IsValid)
                    {
                        pendingItem.URN = oSalesInvoiceInstrument.PostingReference.URN;
                        pendingItem.Post();
                        if (pendingItem.SiJcTrn != null)
                        {
                            //update Project transaction with database keys after posting
                            pendingItem.SiJcTrn.SLPostedCustomerTranDbKey = oSalesInvoiceInstrument.ActualPostedAccountEntryID;
                            pendingItem.SiJcTrn.Trans = oSalesInvoiceInstrument.ActualPostedAccountEntryID;
                            pendingItem.SiJcTrn.SLCustomerAccountDbKey = oSalesInvoiceInstrument.Trader.TradingAccountID;
                            pendingItem.SiJcTrn.Update();
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}
