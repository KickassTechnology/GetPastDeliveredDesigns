using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MsgReader.Exceptions;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;


namespace GetPastDeliveredDesigns
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Extract the email folder to a path where you can use it
            String folderPath = "C:\\Users\\Joe\\Desktop\\DistributionEmails\\Files";

            Int32 iCountNum = 0;

            //Only open files that are email messages
                foreach (String file in Directory.EnumerateFiles(folderPath, "*.msg"))
                {
                    try
                    {
                        OutlookStorage.Message outlookMsg = new OutlookStorage.Message(file);

                    //Loop through the messages attachments
                    foreach (OutlookStorage.Attachment attach in outlookMsg.Attachments)
                    {
                        
                        byte[] attachBytes = attach.Data;

                        if (attach.Filename.IndexOf(".pdf") > -1 || attach.Filename.IndexOf(".dwf") > -1)
                        {

                            textBox1.AppendText(iCountNum.ToString() + ". " + file + System.Environment.NewLine);
                            int iPos = file.LastIndexOf("\\") + 1;
                            
                            //The filenames come in a few formats. This is the best we have to information about what's in the email
                            string filename = file.Substring(iPos, (file.Length - iPos));
                            String[] arrItems = Regex.Split(filename, " - ");

                            //connect to the database
                            String strConn = "Server=mlawdb.cja22lachoyz.us-west-2.rds.amazonaws.com;Database=MLAW_MS;User Id=sa;Password=!sd2the2power!;";
                            SqlConnection conn = new SqlConnection(strConn);

                            String strMLAWNumber = "";

                            if (arrItems.Length > 1)
                            {
                                //First item is always the date, strip out the word Transmittal if it's in there
                                String strDate = arrItems[0];
                                strDate = strDate.Replace("Transmittal", "").Trim();
                                DateTime dtTime = DateTime.MinValue;

                                try
                                {
                                    dtTime = Convert.ToDateTime(strDate);

                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(file + ":::" + strDate);
                                }


                                textBox1.AppendText(dtTime.ToString() + System.Environment.NewLine);

                                for (int i = 1; i < arrItems.Length; i++)
                                {

                                    textBox1.AppendText(arrItems[i].ToString() + System.Environment.NewLine);

                                    if (!isTransmittal(arrItems[i].ToString()))
                                    {

                                        String strThis = arrItems[i].ToString().Trim();
                                        strThis = strThis.Replace(".msg", "");

                                        int iCount = strThis.Count(char.IsLetter);

                                        if (iCount < 2 && strThis.Length > 7)
                                        {
                                            iCountNum += 1;

                                            strThis = strThis.Replace(" ", ".");

                                            strMLAWNumber = strThis;


                                            textBox1.AppendText(iCountNum.ToString() + ". " + strThis + System.Environment.NewLine);
                                        }
                                    }
                                }


                                if (strMLAWNumber == "")
                                {
                                    String strAddress = "";

                                    for (int i = 1; i < arrItems.Length; i++)
                                    {

                                        if (!isTransmittal(arrItems[i].ToString()))
                                        {

                                            String strThis = arrItems[i].ToString().Trim();
                                            strThis = strThis.Replace(".msg", "");

                                            int iCount = strThis.Count(char.IsLetter);
                                            textBox1.AppendText("Address: " + strThis + System.Environment.NewLine);


                                            DataSet ds = new DataSet();

                                            //Try to find our MLAW Number based on the Address and the date delivered
                                            SqlCommand sqlComm = new SqlCommand("Get_MLAW_Number_By_Delivery", conn);
                                            sqlComm.Parameters.AddWithValue("@Address", strThis);
                                            sqlComm.Parameters.AddWithValue("@Date", dtTime);

                                            sqlComm.CommandType = CommandType.StoredProcedure;

                                            SqlDataAdapter da = new SqlDataAdapter();
                                            da.SelectCommand = sqlComm;

                                            da.Fill(ds);

                                            if (ds.Tables[0].Rows.Count > 0 && strMLAWNumber == "")
                                            {
                                                strMLAWNumber = ds.Tables[0].Rows[0]["MLAW_Number"].ToString();

                                            }

                                        }
                                    }

                                    //If we manage to get an MLAW Number then we can do something with our file.
                                    if (strMLAWNumber != "")
                                    {
                                        DataSet dsMLAWNum = new DataSet();

                                        SqlCommand sqlCommMLAWNum = new SqlCommand("Get_Order_By_MLAW_Number", conn);
                                        sqlCommMLAWNum.Parameters.AddWithValue("@MLAW_Number", strMLAWNumber);

                                        sqlCommMLAWNum.CommandType = CommandType.StoredProcedure;

                                        SqlDataAdapter daMLAWNum = new SqlDataAdapter();
                                        daMLAWNum.SelectCommand = sqlCommMLAWNum;

                                        daMLAWNum.Fill(dsMLAWNum);

                                        if (dsMLAWNum.Tables[0].Rows.Count > 0)
                                        {
                                            SqlParameter fileP = new SqlParameter("@file", SqlDbType.VarBinary);
                                            fileP.Value = attachBytes;

                                            SqlParameter sqlOrderId = new SqlParameter("@MLAW_Number", SqlDbType.VarChar);
                                            sqlOrderId.Value = strMLAWNumber;

                                            SqlParameter sqlFileName = new SqlParameter("@File_Name", SqlDbType.VarChar);
                                            String strFileInDB = attach.Filename;

                                            sqlFileName.Value = strFileInDB;

                                            SqlCommand myCommand = new SqlCommand();
                                            myCommand.Parameters.Add(fileP);
                                            myCommand.Parameters.Add(sqlOrderId);
                                            myCommand.Parameters.Add(sqlFileName);


                                            conn.Open();
                                            myCommand.Connection = conn;
                                            myCommand.CommandText = "Insert_Order_File_2";
                                            myCommand.CommandType = CommandType.StoredProcedure;
                                            myCommand.ExecuteNonQuery();
                                            conn.Close();
                                        }
                                    }
                                }
                            }
                        }

                    }
                        if (outlookMsg != null)
                        {
                            outlookMsg.Dispose();
                        }
                    }
                    catch(System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
             
                }

        }


        //Transmittal gets spelled a bunch of different ways. Take care of all of them.
        public bool isTransmittal(String strText)
        {
            bool bIsTranmittal = false;

            if (strText.ToLower().Trim() == "transmital" ||
                strText.ToLower().Trim() == "transmittal" ||
                strText.ToLower().Trim() == "transmittla" ||
                strText.ToLower().Trim() == "transmitla" ||
                strText.ToLower().Trim() == "transittal" ||
                strText.ToLower().Trim() == "tranmittal"
                )
            {

                bIsTranmittal = true;
            }
            return (bIsTranmittal);
        }
    }
}
