using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace solidworks.propreader
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //============================================================================================================
                //Connect with running SolidWorks
                //============================================================================================================
                SldWorks.SldWorks cadRef = (SldWorks.SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                //============================================================================================================
                //Display connected version
                //============================================================================================================
                Console.WriteLine(string.Format("You are connected to SolidWorks {0}", cadRef.RevisionNumber()));
                //============================================================================================================
                //Connect to active document
                //============================================================================================================
                SldWorks.ModelDoc2 cadDocument = cadRef.ActiveDoc;
                Console.WriteLine(string.Format("You are connected to SolidWorks File {0}", cadDocument.GetPathName()));
                //============================================================================================================
                //Read properties
                //============================================================================================================
                SldWorks.CustomPropertyManager managerCustom = null;
                object propertyNames = null;
                object propertyTypes = null;
                object propertyValues = null;
                object propertyResolvedValues = null;
                int propertyCount = 0;
                //Start with custom properties
                Console.WriteLine("====================================CUSTOM=============================================================");
                managerCustom = cadDocument.Extension.CustomPropertyManager[string.Empty];
                propertyCount = managerCustom.GetAll2(ref propertyNames, ref propertyTypes, ref propertyValues, ref propertyResolvedValues);
                for (int i = 0; i < propertyCount; i++)
                {
                    string valOut = "";
                    string resolvedValOut = "";
                    bool wasResolved = false;
                    int result = managerCustom.Get5(((Array)propertyNames).GetValue(i).ToString(), false, out valOut, out resolvedValOut, out wasResolved);
                    Console.WriteLine(string.Format("Property Name - {0}, Property Type - {1}, Property Value - {2}, Property Resolved Value - {3}"
                        , ((Array)propertyNames).GetValue(i).ToString()
                        , ((SwConst.swCustomInfoType_e)((Array)propertyTypes).GetValue(i)).ToString()
                        , valOut
                        , resolvedValOut));
                }
                Console.WriteLine("========================================================================================================");
                //Now read configuration specific properties
                //First get all configurations
                object configNames = null;
                configNames = cadDocument.GetConfigurationNames();
                foreach (var config in (Array)configNames)
                {
                    Console.WriteLine(string.Format("===================================={0}=============================================================",config.ToString()));
                    managerCustom = cadDocument.Extension.CustomPropertyManager[config.ToString()];
                    propertyCount = managerCustom.GetAll2(ref propertyNames, ref propertyTypes, ref propertyValues, ref propertyResolvedValues);
                    for (int i = 0; i < propertyCount; i++)
                    {
                        string valOut = "";
                        string resolvedValOut = "";
                        bool wasResolved = false;
                        int result = managerCustom.Get5(((Array)propertyNames).GetValue(i).ToString(), false, out valOut, out resolvedValOut, out wasResolved);
                        Console.WriteLine(string.Format("Property Name - {0}, Property Type - {1}, Property Value - {2}, Property Resolved Value - {3}"
                            , ((Array)propertyNames).GetValue(i).ToString()
                            , ((SwConst.swCustomInfoType_e)((Array)propertyTypes).GetValue(i)).ToString()
                            , valOut
                            , resolvedValOut));
                    }
                    Console.WriteLine("========================================================================================================");
                }
                //============================================================================================================
                //Read Summary Information
                //============================================================================================================
                var author = cadDocument.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoAuthor];
                var keyWords = cadDocument.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoKeywords];
                var comments = cadDocument.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoComment];
                var title = cadDocument.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoTitle];
                var subject = cadDocument.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoSubject];
                Console.WriteLine("====================================SUMMARY=============================================================");
                Console.WriteLine(string.Format("Author - {0}", author));
                Console.WriteLine(string.Format("Keywords - {0}", keyWords));
                Console.WriteLine(string.Format("Comments - {0}", comments));
                Console.WriteLine(string.Format("Title - {0}", title));
                Console.WriteLine(string.Format("Subject - {0}", subject));                
                Console.WriteLine("========================================================================================================");
                //============================================================================================================
                //Ask and close
                //============================================================================================================
                Console.WriteLine("Press enter to close...");
                Console.ReadLine();
                //============================================================================================================
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Press enter to close...");
                Console.ReadLine();
            }
        }
    }
}
