using Sicon.Sage200.Projects.Objects;
using Sicon.Sage200.Projects.Objects.Infrastructure.Projects.Factories;
using System;

namespace ProjectsExamples
{
    public class ProjectMethods
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
        /// Get a project by its ID
        /// </summary>
        /// <param name="SiJcJobID"></param>
        /// <returns></returns>
        public SiJcJob GetProject(long SiJcJobID)
        {
            try
            {
                //Get project by project ID
                return ProjectFactory.Factory.Fetch(SiJcJobID);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Get a project by its Number
        /// </summary>
        /// <param name="ProjectNumber">For Example 'P00001'</param>
        /// <returns></returns>
        public SiJcJob GetProject(string ProjectNumber)
        {
            try
            {
                //Get project by project number
                return ProjectFactory.Factory.FetchWithProjectNumber(ProjectNumber);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Method to create or amend a project
        /// </summary>
        /// <param name="ProjectNumber"></param>
        /// <param name="StartDate"></param>
        /// <param name="PlannedCompletionDate"></param>
        /// <param name="ActualCompletionDate"></param>
        /// <param name="PercentageComplete"></param>
        /// <param name="PercentageStockCover"></param>
        /// <param name="oCustomer"></param>
        /// <param name="CustomerDeliveryLocationRef"></param>
        /// <param name="AnalysisCode1"></param>
        /// <param name="AnalysisCode2"></param>
        /// <param name="AnalysisCode3"></param>
        /// <param name="AnalysisCode4"></param>
        /// <param name="AnalysisCode5"></param>
        /// <param name="AnalysisCode6"></param>
        /// <param name="AnalysisCode7"></param>
        /// <param name="DefaultCostCentre"></param>
        /// <param name="DefaultDepartmentCode"></param>
        public void CreateOrAmendProject(string ProjectNumber, DateTime StartDate, DateTime PlannedCompletionDate, DateTime ActualCompletionDate, decimal PercentageComplete, decimal PercentageStockCover, Sage.Accounting.SalesLedger.Customer oCustomer, 
            string CustomerDeliveryLocationRef, string AnalysisCode1, string AnalysisCode2, string AnalysisCode3, string AnalysisCode4, string AnalysisCode5, string AnalysisCode6, string AnalysisCode7, string DefaultCostCentre, string DefaultDepartmentCode)
        {
            try
            {
                //Get project by project ID
                SiJcJob oProject = ProjectFactory.Factory.FetchWithProjectNumber(ProjectNumber);
                if (oProject == null)
                {
                    //Create new Project
                    oProject = new SiJcJob();
                    oProject.AllJobHeadersAvailable = true;
                    oProject.JobNumber = ProjectNumber;
                    oProject.LastC2CPeriod = 0;
                    oProject.LastC2CByUser = "";
                    oProject.LastC2CDateTime = DateTime.MinValue;
                    oProject.SYSCurrency = Sage.Accounting.SystemManager.FinancialCurrencyFactory.Factory.GetBaseCurrency().SYSCurrencyID;
                }

                //Amend Project
                oProject.Description = "Project Description";
                oProject.JobManager = "Joe Bloggs";
                oProject.StartDate = StartDate.Date;
                oProject.PlanCompDate = PlannedCompletionDate.Date;
                oProject.ActualCompDate = ActualCompletionDate.Date;
                oProject.PercentageComplete = PercentageComplete;
                oProject.PercentStockCover = PercentageStockCover;
                oProject.ChargeableType = 0;//
                if (oProject.PercentageComplete < 100)
                {
                    oProject.JCStatus = "Live";
                }
                if (oProject.PercentageComplete == 100)
                {
                    oProject.JCStatus = "Complete";
                }

                //Set With Sage customer detail
                if (oCustomer != null)
                {
                    oProject.TraderAccount = oCustomer.TradingAccountID;
                    oProject.CustomerName = oCustomer.Name;

                    //Set Customer Address
                    if (CustomerDeliveryLocationRef != "")
                    {
                        Sage.Accounting.SOP.CustDeliveryAddresses oDelAdds = new Sage.Accounting.SOP.CustDeliveryAddresses();
                        oDelAdds.Query.Filters.Add(new Sage.ObjectStore.Filter(Sage.Accounting.SOP.CustDeliveryAddress.FIELD_CUSTOMERDBKEY, oCustomer.SLCustomerAccount));
                        oDelAdds.Query.Filters.Add(new Sage.ObjectStore.Filter(Sage.Accounting.SOP.CustDeliveryAddress.FIELD_DESCRIPTION, CustomerDeliveryLocationRef));
                        oDelAdds.Find();

                        if (!oDelAdds.IsEmpty)
                        {
                            oProject.CustDeliveryAddress = oDelAdds.First.CustDeliveryAddress;
                            oProject.AddressPostalName = oDelAdds.First.PostalName;
                            oProject.AddressLine1 = oDelAdds.First.AddressLine1;
                            oProject.AddressLine2 = oDelAdds.First.AddressLine2;
                            oProject.AddressLine3 = oDelAdds.First.AddressLine3;
                            oProject.AddressLine4 = oDelAdds.First.AddressLine4;
                            oProject.AddressCity = oDelAdds.First.City;
                            oProject.AddressPostCode = oDelAdds.First.PostCode;
                            oProject.AddressContact = oDelAdds.First.Contact;
                            oProject.AddressTelephone = oDelAdds.First.TelephoneNo;
                            oProject.AddressFax = oDelAdds.First.FaxNo;
                            oProject.AddressEmail = oDelAdds.First.EmailAddress;
                        }
                    }
                }

                //Analysis Codes
                oProject.AnalysisType1 = AnalysisCode1;
                oProject.AnalysisType2 = AnalysisCode2;
                oProject.AnalysisType3 = AnalysisCode3;
                oProject.AnalysisType4 = AnalysisCode4;
                oProject.AnalysisType5 = AnalysisCode5;
                oProject.AnalysisType6 = AnalysisCode6;
                oProject.AnalysisType7 = AnalysisCode7;

                //Setting Project nominal overrides
                if (DefaultCostCentre != "")
                {
                    oProject.JobCostCentreOverride = true;
                    oProject.JobCostCentreOverrideCode = DefaultCostCentre;
                }
                else
                {
                    oProject.JobCostCentreOverride = false;
                }
                if (DefaultDepartmentCode != "")
                {
                    oProject.JobDepartmentOverride = true;
                    oProject.JobDepartmentOverrideCode = DefaultDepartmentCode;
                }
                else
                {
                    oProject.JobDepartmentOverride = false;
                }

                //Create default variation if it doesnt exist
                if (oProject.SiJcVariations.IsEmpty)
                {
                    SiJcVariation oSiJcVariation = (SiJcVariation)oProject.SiJcVariations.AddNew();
                    oSiJcVariation.Description = "Variation0";
                    oSiJcVariation.DateAdded = Sicon.API.Sage200.Objects.StaticClass.GetCurrentDateTime();
                    oSiJcVariation.SiJcJobDbKey = oProject.SiJcJob;
                    oSiJcVariation.Update();
                }

                //Mark as not deleted if this job was previously deleted.
                oProject.Deleted = "N";
                oProject.DeletedUser = "";
                oProject.DeletedDate = DateTime.MinValue;
                oProject.UpdatedDate = DateTime.Today;
                oProject.UpdatedUser = Sicon.API.Sage200.Objects.StaticClass.GetSageActiveUserName();

                //Commit project
                oProject.Update();

            }
            catch (Exception)
            {

                throw;
            }
        }
    
    
        /// <summary>
        /// Add Links to a project
        /// </summary>
        /// <param name="oProject"></param>
        public void CreateProjectHeaderLinksOnAProject(SiJcJob oProject)
        {
            try
            {
                SiJcChds oSiJcChds = new SiJcChds();
                oSiJcChds.Query.Filters.Add(new Sage.ObjectStore.Filter(SiJcChd.FIELD_DELETEDUSER, Sage.Common.Data.FilterOperator.Equal, ""));
                oSiJcChds.Query.Filters.Add(new Sage.ObjectStore.Filter(SiJcChd.FIELD_INACTIVE, Sage.Common.Data.FilterOperator.Equal, false));
                oSiJcChds.Query.Sorts.Add(new Sage.ObjectStore.Sort(SiJcChd.FIELD_COSTCODE, true));
                oSiJcChds.Find();

                foreach (SiJcChd oSiJcChd in oSiJcChds)
                {
                    //Enable job header on job
                    SiJcJrt oSiJcJrt = ProjectHeaderLinkFactory.Factory.Fetch(oProject.SiJcJob, oSiJcChd.SiJcChd, true);
                    if (oSiJcJrt == null)
                    {
                        oSiJcJrt = new SiJcJrt();
                        oSiJcJrt.ResourcePercentageMarkup = 0;
                        oSiJcJrt.SiJcJobDbKey = oProject.SiJcJob;
                        oSiJcJrt.SiJcChdDbKey = oSiJcChd.SiJcChd;
                        oSiJcJrt.ChargeType = "C";
                        oSiJcJrt.LastC2CByUser = "";
                        oSiJcJrt.LastC2CDateTime = DateTime.MinValue;
                        oSiJcJrt.LastC2CPeriod = 0;
                        oSiJcJrt.UpdatedDate = Sage.Common.Clock.Now;
                        oSiJcJrt.UpdatedUser = Sicon.API.Sage200.Objects.StaticClass.GetSageActiveUserName();
                        oSiJcJrt.Available = true;
                        oSiJcJrt.Update();
                    }
                    else
                    {
                        oSiJcJrt.Available = true;
                        oSiJcJrt.Update();
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
