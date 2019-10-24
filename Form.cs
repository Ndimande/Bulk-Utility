
using Newtonsoft.Json;
using RestClientTest;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RESTAPITest1
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            
            InitializeComponent();
           
            string endpoint =Endpoint.Text;
            

        }

      
       
        Login login = new Login();
        RestTest restClient = new RestTest();
      
        

        Read_From_Excel accessExcel = new Read_From_Excel();
       


        private void button1_Click(object sender, EventArgs e)
        {
           
           

        }

        private void txtuserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        //private void moveFile_Click(object sender, EventArgs e)
        //{
        //    //eg moveNode("25347", "25453");
        //    this.timer1.Start();
        //    txtResponse.Text = restClient.moveNode(MoveNodeID.Text, MoveNodeIDParentID.Text, restClient.Authenticate(login.username(), login.password()));
        //}

        //private void createFolder_Click(object sender, EventArgs e)
        //{
        //    this.timer1.Start();
        //    txtResponse.Text = restClient.createNewFolder(newFolderName.Text, FolderparentID.Text, 0, restClient.Authenticate(login.username(), login.password()));

        //}

  

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void main_Click(object sender, EventArgs e)
        {
        }

        private string openFileDialog1_FileOk()
        {

            OpenFileDialog openDialog = new OpenFileDialog();
            string file = openDialog.FileName;
            openDialog.Title = "Select A File";
            openDialog.Filter = "All Files (*.*)|*.*" + "|" +
                                "Text Files (*.txt)|*.txt" + "|" +
                                "Image Files (*.png;*.jpg)|*.png;*.jpg";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                file = openDialog.InitialDirectory + openFileDialog1.FileName; ;


            }
            return file;
        }

        //private void button4_Click(object sender, EventArgs e)
        //{
            
        //    //Path.GetDirectoryName(filePath)
        //    this.timer1.Start();
        //    string Location = @"C:\Users\pndimande\Desktop\Travel Request Template.docx";
        //    // @"C:\Users\pndimande\Desktop\Travel Request Template.docx


        //    nameFile.Text = Location.Substring(Location.LastIndexOf('\\') +
        //     1);

            


        //    txtResponse.Text = restClient.uploadDocument(openFileDialog1_FileOk(), nameFile.Text, nameFile.Text.Substring(nameFile.Text.Length - 5),
        //        ParentID.Text, 144, restClient.Authenticate(login.username(), login.password())).ToString();



        //} // If you want to upload a single file from c drive

        private void nodeID_TextChanged(object sender, EventArgs e)
        {

        }


        /*
          
             private void txtNodeID_Click(object sender, EventArgs e)
        {
            Read_From_Excel obj = new Read_From_Excel();

            //   foreach(string item in obj.nodeId(@"C:\Users\pndimande\Desktop\Test.xlsx"))
            //  {

            //Getting the file name from the location
            string Location = @"C:\Users\pndimande\Desktop\Travel Request Template.pdf";

            nameFile.Text = Location.Substring(Location.LastIndexOf('\\') +
                1);

            txtPath.Text = @"C:\Users\pndimande\Desktop\Test.xlsx";
            string[] array = obj.nodeId(txtPath.Text);
            int length = array.Length -1;
            for (int i  = 0; i<= length; i++)
                {



                if (restClient.search(array[i]) == "File Found")
                {
                    
                    restClient.uploadDocument(@"C:\Users\pndimande\Desktop\Travel Request Template.pdf", nameFile.Text, nameFile.Text.Substring(nameFile.Text.Length - 5), array[i], 144).ToString();
                    nodeID.Text = obj.nodeId(txtPath.Text)[i];
                }

               


            }
                
        }
          
       *///Method below copied

        private void txtNodeID_Click(object sender, EventArgs e)
        {

            this.timer1.Start();
            int count = 0;
            int counter = 1;
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Test2.xlsx";
            txtPath.Text = Excel_To_Be_Read;
            string[] arrayID = accessExcel.nodeId(txtPath.Text);
            string[] arrayName = accessExcel.getExcelFile(txtPath.Text);
            int lengthID = arrayID.Length - 1;
            for (int i = 0; i <= lengthID; i++)
            {


                txtFullFileName.Text = arrayName[i];
                txtBWS.Text = accessExcel.BWS(txtPath.Text)[i];
                nodeID.Text = arrayID[i];
                //Printing the the first four chars wgich is the business workspance name.
                //  string temp = txtFullFileName.Text;
                // txtBWS.Text = temp;

                var searchNode = restClient.search(arrayID[i], ticket);


                if (searchNode == "File Found")
                {
                    //id -parent_ID

                    restClient.moveNode(arrayID[i], restClient.ReturnNodeIDSearch(txtBWS.Text,ticket), ticket);

                    debugOutput( "Sucessfully moved NO: " + counter +": "+ arrayName[i] + "\n");
                    count++;
                    countLabel.Text = count.ToString();
                    counter++;
                }
                
               



            }

        }



        private void txtPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            string ticket = restClient.Authenticate(login.username(), login.password());

            string attributeIDString = string.Format("Attr_{0} : {1}", txtAttributeID.Text, txtValueSearch.Text);

            restClient.search(attributeIDString, txtNewValue.Text, txtResponse, txtAttributeID.Text, ticket, countLabel);
        }



        private void CreateBWS_Click(object sender, EventArgs e)
        {
            this.timer1.Start();
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 0;
            int errorCount = 0;
            string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Test4.xlsx";
            ExtelFilePath.Text = Excel_To_Be_Read;
            string[] BWS = accessExcel.BWSName(ExtelFilePath.Text);
            string[] arrayCompanyName = accessExcel.siteName(ExtelFilePath.Text);
            string id = string.Empty;
            int lengthID = BWS.Length - 1;
            string[] array = new string[lengthID+1];
            for (int i = 2; i <= lengthID; i++)
            {

                //first  2 values returns null that why there is bad request

               // txtResponse.Text = arrayCompanyName[2] + "\t"+ BWS[2] +"\n";

                CompanyName.Text = arrayCompanyName[i];
                BWSNAME.Text = BWS[i];

                //Printing the the first four chars wgich is the business workspance name.
                // string temp = txtFullFileName.Text;
                // txtBWS.Text = temp;

                // var searchNode = restClient.search(arrayCompanyName[i], restClient.Authenticate(login.username(), login.password()));


                // if (searchNode == "File Found")
                //{
                //id -parent_ID
                // string template_id, string parent_id, string workspaceName ,string ticket

 

               
                TemplateIDBWS.Text = "66610";
                ParentIDBWS.Text = "66395";
                int counter = 1;
                
                //Im getting a bad request for the first bws creation
                string bwsCreated = restClient.createBusinessWorkspace(TemplateIDBWS.Text, ParentIDBWS.Text, BWS[i] + "\t" + arrayCompanyName[i], ticket) ;
                // txtResponse.Text = bwsCreated;
                if (bwsCreated == "Error: File already exist in Content server or the node ID was wrong!")
                {
                    count--;
                    errorCount++;
                }
                array[i] =bwsCreated.ToString();

                countLabel.Text =count.ToString();
                error.Text = errorCount.ToString();
                debugOutput(bwsCreated  +" No: "+ counter);
                count++;

            }

        }

        private void debugOutput(string strDebugText)
        {
            try
            {
                System.Diagnostics.Debug.Write(strDebugText + Environment.NewLine);
                txtResponse.Text = txtResponse.Text + strDebugText + Environment.NewLine;
                txtResponse.SelectionStart = txtResponse.TextLength;
                txtResponse.ScrollToCaret();
            }

            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex.Message, ToString() + Environment.NewLine);
            }
        }


        private void TestButton_Click(object sender, EventArgs e)
        {
            ////  getCategoryIDAndCategoryAttributeID(string ticket, string Catagory_ID)

            // updateAttribute();

            //var nodeID =restClient.ReturnNodeIDSearch("Test10111", ticket);
            //var categoryIDInJson = restClient.getCategoryIDAndCategoryAttributeID(nodeID, ticket);
            ////public void InsertAttributeValues(string ticket, int id, string changeValue, TextBox txtResponse, string attributeID, int i, string endpoint)
            ////restClient.InsertAttributeValues();

            //string attributeIDString = string.Format("Attr_{0} : {1}", txtAttributeID.Text, txtValueSearch.Text);

            //restClient.search(attributeIDString, txtNewValue.Text, txtResponse, txtAttributeID.Text, ticket, countLabel);
            ////var ancestors = restClient.GetAncestors(search, ticket);

            // txtResponse.Text = restClient.getCategoryIDAndCategoryAttributeID(ticket ,); //read the name of a file and search for it then returns the node node id of it

            this.timer1.Start();
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();

            string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Test4.xlsx";
            ExtelFilePath.Text = Excel_To_Be_Read;
            string[] BWS = accessExcel.BWSName(ExtelFilePath.Text);
            
           
            string[] arrayCompanyName = accessExcel.siteName(ExtelFilePath.Text);


            string id = string.Empty;
           // id = restClient.ReturnNodeIDSearch("7FNW" + "\t" + "NOORD MOTORS", ticket);
           // restClient.insertValues(ticket, id, "fawazz", "pATRICKDDJFN");
            //txtResponse.Text = arrayCompanyName[];
            int lengthID = BWS.Length - 1;
            int counter = 1;
           
            for (int i = 2; i <= lengthID; i++)
            {

                
                //BWS[temp]
                //arrayCompanyName[temp]
                string item1 = BWS[i];
                string item2 = arrayCompanyName[i];
                //txtResponse.Text = item1 + "\t" + item2;
                id = restClient.ReturnNodeIDSearch(item1 + "\t" + item2 , ticket);

                restClient.insertValues(ticket, id, item1, item2);

                debugOutput("Catergory values has been inserted " + counter);
              
                counter++;

            }
            


        }

        private void txtValueSearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtAttributeID_TextChanged(object sender, EventArgs e)
        {

        }

     

      

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(1);
        }

        private void nameFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 0;
            int errorCount = 0;
            string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Test5.xlsx";
            Path.Text = Excel_To_Be_Read;
            string[] parentID = accessExcel.parentID(Path.Text);//This is fine 
            string[] recordType = accessExcel.recordType(Path.Text); //This is fine
            string[] path = accessExcel.filePath(Path.Text); //This is fine
            string[] subfolder = accessExcel.subFolder(Path.Text); //This is fine
           // string[] categoryID = accessExcel.categoryID(Path.Text); //This is fine
           if("String".ToString() == "String")
            {
                
            }
          
            int lengthID = parentID.Length - 1;
           // string[] array = new string[lengthID + 1];
       
            this.timer1.Start();

            for(int i = 2; i< lengthID; i++)
            {
                string []Location = path;

                names.Text = Location[i].Substring(Location[i].LastIndexOf('\\') +
                 1);
            

                txtResponse.Text = restClient.uploadDocument(Location[i].ToString(), names.Text, names.Text.Substring(names.Text.Length - 5),
                    parentID[i].Trim(), 144, restClient.Authenticate(login.username(), login.password())).ToString();
                count++;
                if (txtResponse.Text == "There was a problem uploading a file, server or that folder does not exist.")
                {
                    count--;
                    errorCount++;
                }
                //Category vlue 26007_5
                error.Text = errorCount.ToString();
                countLabel.Text = count.ToString();

               
            }



           
        }

        private void Update_Category_Click(object sender, EventArgs e)
        {
            int count = 1;
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            string id = string.Empty;
            string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Test5.xlsx";
            Path.Text = Excel_To_Be_Read;
            string[] parentID = accessExcel.parentID(Path.Text);//This is fine 
            string[] recordType = accessExcel.recordType(Path.Text); //This is fine
            string[] path = accessExcel.filePath(Path.Text); //This is fine
            string[] subfolder = accessExcel.subFolder(Path.Text); //This is fine
            string[] categoryID = accessExcel.categoryID(Path.Text); //This is fine


            int lengthID = parentID.Length - 1;
            for (int i = 2; i < lengthID; i++)
            {
                restClient.insertValuesForHR(ticket, categoryID[i], recordType[i]);

                debugOutput("Catergory values has been inserted " + count);
                count++;
            }

        }


        private void createFolders_Click(object sender, EventArgs e)
        {
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 1;
            int errorCount = 1;
            // string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Sample.xlsx";
            //FolderPath.Text = Excel_To_Be_Read;
            if (FolderPath .Text != "")
            {
                string[] parentID = accessExcel.parentIDForFolderCreation(FolderPath.Text);//This is fine 
                string[] FolderName = accessExcel.FolderName(FolderPath.Text); //This is fine

                for (int i = 2; i < parentID.Length; i++)
                {
                    var answer = restClient.createNewFolder(FolderName[i].Trim(), parentID[i].Trim(), 0, ticket);


                    if (answer == "Folder Created")
                    {
                        Cname.Text = FolderName[i].Trim();
                        PIDname.Text = parentID[i].Trim();

                        debugOutput(answer + " No:" + count + " of Parent ID: " + parentID[i].Trim());

                        countLabel.Text = count.ToString();
                        this.timer1.Start();
                        count++;
                    }
                    else
                    {
                        this.timer1.Start();
                        debugOutput("The was a problem when creating the folder No " + errorCount + " of Parent ID: " + parentID[i].Trim());
                        error.Text = errorCount.ToString();
                        errorCount++;
                    }

                }
            }
            else
            {
                MessageBox.Show("Please provide excel path");
            }
        }
        private void rename_Click(object sender, EventArgs e)
        {

            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 1;
            int errorCount = 1;
            //string Excel_To_Be_Read = @"C:\Users\pndimande\Desktop\Sample.xlsx";
            //renameF.Text = Excel_To_Be_Read;
            if (renameF.Text != "")
            {
                string[] NodeID = accessExcel.nodeId(renameF.Text);//This is fine 
                string[] newValue = accessExcel.FolderValues(renameF.Text); //This is fine

                for (int i = 2; i < newValue.Length; i++)
                {

                    var answer = restClient.UpdateFolderName(ticket, NodeID[i].Trim(), newValue[i].Trim());
                    if (answer == "Folder Updated")
                    {
                        PD.Text = NodeID[i];
                        Fname.Text = newValue[i].Trim();
                        debugOutput(answer + " No " + count + " Node ID: " + NodeID[i].Trim());
                        countLabel.Text = count.ToString();
                        count++;


                    }
                    else
                    {
                        debugOutput("The was a problem when creating the folder No " + errorCount + " of Node ID " + NodeID[i].Trim());
                        error.Text = errorCount.ToString();
                        errorCount++;
                        this.timer1.Start();
                    }

                }
            }
            else
            {
                MessageBox.Show("Please provide excel path");
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int NumberOfFolders = 6;
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 0;
            for (int i = 0; i<= NumberOfFolders; i++)
            {
                NodeIDUpdate.Text = restClient.ReturnNodeIDSearch(searchUpdate.Text, ticket);
                debugOutput(restClient.UpdateFolderName(ticket, NodeIDUpdate.Text, newValueUpdate.Text) + " No: " + count + " is Updated") ;
                
                count ++;
                
            }
            countLabel.Text = count.ToString();

            
        }

        private void BTPermissions_Click(object sender, EventArgs e)
        {
            string ticket = restClient.Authenticate(login.username(), login.password()).ToString();
            int count = 1;
            int errorCount = 1;
            string nodeID = "3374811";
            string right_id = "1508399";
           debugOutput(restClient.Permissions(nodeID,right_id,ticket));
        }
    }
}

