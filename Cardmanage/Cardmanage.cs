
using System;
using HS.Platform;
using System.IO;
using System.Data;

namespace Cardmanage
{

    public class Cardmanage :HS.Platform.IRecordEditCustomize   
    {
        public string filefullname = "D:\\BlackRecord\\WaterControlBlackList.csv";
        public string filedir = "D:\\BlackRecord\\";
        public string filefullnames = "D:\\BlackRecordServer\\WaterControlBlackList.csv";
        public string filedirs = "D:\\BlackRecordServer\\";
        public void DealAfterEdit(ref CmsPassport pst, long lngResID, long lngRecID, ref System.Collections.Hashtable hashFieldVal, ref System.Collections.Hashtable hashOldFieldVal, ResSyncAction intSyncAction)
        {
            //处理挂失文件
           
                if (!Directory.Exists(filedir))
                {
                    Directory.CreateDirectory(filedir);
                }
                if (!Directory.Exists(filedirs))
                {
                    Directory.CreateDirectory(filedirs);
                }
            
                //if (intSyncAction ==ResSyncAction.AddRecord)
                //{
                   
                //    string cardno = Convert.ToString(hashFieldVal["C3_309560350826"]);
                //    string cardstatus = Convert.ToString(hashFieldVal["C3_374115188913"]);
                //    if (cardstatus != "正常")
                //    {
                //        StreamWriter sw = new StreamWriter(filefullname, true);
                       
                //        sw.WriteLine(cardno);
                //        sw.Close();  
                //    }
                 
                //}
                if (intSyncAction == ResSyncAction.EditRecord)
                {


                    string cardno = "";//; = Convert.ToString(hashFieldVal["C3_309560350826"]);
                  
                        StreamWriter sw = new StreamWriter(filefullname, false);
                        
                    
                        DataSet dt = new DataSet();
                        string strSql = "select * from  CT301050266340 where C3_301064278658<>1";
                        Int32 intRecTotalAmount = 0;
                        dt=CmsTable.GetDatasetForHostTable(ref pst, 301050266340, false, "", "", strSql, 0, 0, ref intRecTotalAmount, "", false);
                        for (int i = 0; i < dt.Tables[0].DefaultView.Count; i++)

                        {
                            cardno = Convert.ToString(dt.Tables[0].Rows[i]["C3_301064188869"]);
                            sw.WriteLine(cardno);
                            
                        }
                        sw.Close();

                        StreamWriter sw1 = new StreamWriter(filefullnames, false);

                      
                        for (int ii = 0; ii < dt.Tables[0].DefaultView.Count; ii++)
                        {
                            cardno = Convert.ToString(dt.Tables[0].Rows[ii]["C3_301064188869"]);
                            sw1.WriteLine(cardno);

                        }
                        sw1.Close();
                         
                    //}
                   
                  
                  
                }

            
          
        }

        public void DealBeforeEdit(ref CmsPassport pst, long lngResID, long lngRecID, ResSyncAction intSyncAction, ref System.Collections.Hashtable hashFieldVal)
        {
            throw new NotImplementedException();
        }
    }
}
