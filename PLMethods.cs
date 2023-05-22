using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Factory;
using Sicon.Sage200.Projects.Objects.Instruments.PL;
using System;

namespace ProjectsExamples
{
    public class PLMethods
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
        /// <param name="PLPostedTransactionID"></param>
        public void GetProjectTransactionsForPLInvoice(long PLPostedTransactionID)
        {
            try
            {
                //Get project transactions by <long> PLPostedTransactionID
                SiJcTrns oPL_SIJCTRNs = ProjectTransactionsFactory.Factory.FetchPL(PLPostedTransactionID, "PLINV", true);
            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// Post And Reconcile PL Invoice To Project
        /// </summary>
        public void PostAndReconcilePLInvoiceToProject()
        {
            PostPLTransactionInstrument oInstrument = null;
            PLNominalInstrumentItem oNLItem = null;
            PLTaxInstrumentItem oTaxItem = null;
            try
            {
                //Initiate Instrument
                oInstrument = new PostPLTransactionInstrument();
                oInstrument.DocumentType = PostPLTransactionInstrument.DocumentTypeEnum.Invoice;
                oInstrument.Reference = "Test1";
                oInstrument.SecondReference = "Test1";
                oInstrument.SupplierReference = "ATL001";
                oInstrument.TransactionDate = DateTime.Today;

                //Add Nominal Lines
                oNLItem = new PLNominalInstrumentItem();
                oNLItem.NLRef = "03100";
                oNLItem.NLCC = "";
                oNLItem.NLDept = "";
                oNLItem.Narrative = "PI / ATL001 / Test1 / Line1";
                oNLItem.GoodsValue = 50;
                oNLItem.JobNumber = "J0000000001";
                oNLItem.JobHeader = "Material";
                oNLItem.Phase = "";//Leave blank if not using Phases and Stages
                oNLItem.Stage = "";// ''
                oNLItem.Activity = "";// ''
                oInstrument.PLNominalInstrumentItems.Add(oNLItem);

                oNLItem = new PLNominalInstrumentItem();
                oNLItem.NLRef = "02100";
                oNLItem.NLCC = "";
                oNLItem.NLDept = "";
                oNLItem.Narrative = "PI / ATL001 / Test1 / Line2";
                oNLItem.GoodsValue = 50;
                oNLItem.JobNumber = "J0000000001";
                oNLItem.JobHeader = "Labour";
                oNLItem.Phase = "";//Leave blank if not using Phases and Stages
                oNLItem.Stage = "";// ''
                oNLItem.Activity = "";// ''
                oInstrument.PLNominalInstrumentItems.Add(oNLItem);

                //Add Tax Analysis
                oTaxItem = new PLTaxInstrumentItem();
                oTaxItem.TaxCode = 1;
                oTaxItem.GoodsValue = 100;
                oTaxItem.TaxValue = 20;
                oInstrument.PLTaxInstrumentItems.Add(oTaxItem);

                //Post instrument returning URN
                long URN = oInstrument.Post();

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
