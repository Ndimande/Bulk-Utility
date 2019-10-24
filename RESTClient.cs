using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Net.Http;
using RestSharp;
using RestSharp.Serialization.Json;
using RestClientTest;
using RestSharp.Authenticators;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace RestClientTest
{
    public enum httpVerb
    {
        GET,
        POST,
        PUT,
        DELETE
    }


    public class RestTest 
    {

       
        public string endPoint { get; set; }
        public httpVerb httpMethod { get; set; }
        
        
        public RestTest()
        {
          
             //endPoint ="http://eimcs.dcxeim.local:85/otcs/cs.exe";// Datacentrix Endpoint

           // endPoint = "http://cmsqa.engenoil.net/ecm/llisapi.dll";  //QA environment for Engen

            endPoint = "http://ctcsintd01.engenoil.net/ecm/llisapi.dll";  //Dev environment for Engen


        }
      

         

        

        public string Authenticate(string username,string password)
        {
            
            // Creates an empty string
            string ticket = string.Empty;

            // Uses RestSharp as a import and calls it's RestClient method
           
            RestClient client = new RestClient(endPoint);

            // Uses SimpleAuthenticator and adds credentials for authentication- client_id and client_secret
            //client.Authenticator = new SimpleAuthenticator("client_id", "KA5a31zpp41v3Mpn", "client_secret", "2ec0975e36734be297a56dd141f1c98b");
            client.Authenticator = new SimpleAuthenticator("username", username, "password", password);

            // Creates a request call with an added parameter
            var request = new RestRequest("/api/v1/auth", Method.POST);//.AddParameter("grant_type", "client_credentials");


            // Executes the response that you created
            var response = client.Execute(request);

            // This will deserialize the received JSON in a Dictionary - key value pair

            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);
            // In the JSON theres a Key called "access_token" so we grabbing its value
            ticket = output["ticket"];

            // Returns the ticket
            return ticket;
        }


        public string Authentication(string username, string password, string AuthEndPoint)
        {
            // Creates an empty string
            string ticket = string.Empty;

            // Uses RestSharp as a import and calls it's RestClient method

            RestClient client = new RestClient(endPoint);

            // Uses SimpleAuthenticator and adds credentials for authentication- client_id and client_secret
            //client.Authenticator = new SimpleAuthenticator("client_id", "KA5a31zpp41v3Mpn", "client_secret", "2ec0975e36734be297a56dd141f1c98b");
            client.Authenticator = new SimpleAuthenticator("username", username, "password", password);

            // Creates a request call with an added parameter
            var request = new RestRequest("/api/v1/auth", Method.POST);//.AddParameter("grant_type", "client_credentials");


            // Executes the response that you created
            var response = client.Execute(request);

            // This will deserialize the received JSON in a Dictionary - key value pair
            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);

            // In the JSON theres a Key called "access_token" so we grabbing its value
            ticket = output["ticket"];

            // Returns the ticket
            return ticket;
        }//Authentication works

        public string createBusinessWorkspace( string template_id, string parent_id, string workspaceName ,string ticket)
        {
           // string ticket = Authenticate();
            string businessWorkspaceID = string.Empty;

            //Write your code here...............

            RestClient client = new RestClient(endPoint);
            var request = new RestRequest("/api/v2/businessworkspaces", Method.POST);//.AddParameter("grant_type", "client_credentials");
            request.AddParameter("Template_id", template_id);
            request.AddParameter("parent_id", parent_id);
            request.AddParameter("name", workspaceName);
            request.AddHeader("otcsticket", ticket);
            var response = client.Execute(request);
            string status = response.StatusCode.ToString();
            if (status == "OK")
            {
                status = "Business Workspace Created";
            }
            else
            {
                status = "Error: File already exist in Content server or the node ID was wrong!";
            }
            return status;
        }

        public string createNewFolder( string folderName, string parent_id, int type, string ticket)
        {
            string folder_id = string.Empty;
            //string ticket = Authenticate();
            // folderName = "Folderr22";

            RestClient client = new RestClient(endPoint);


            //Example 

            /*
             type = 0
             otcstickets = ticket
             parent_id = 
            */
            var request = new RestRequest("/api/v2/nodes", Method.POST);
            request.AddParameter("type",type);
            request.AddParameter("parent_id", parent_id);
            request.AddParameter("name", folderName);
            request.AddHeader("otcsticket", ticket);

            // Executes the response that you created
            var response = client.Execute(request);

            // This will deserialize the received JSON in a Dictionary - key value pair
            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);
            string status = response.StatusCode.ToString();
            if(status == "OK")
            {
                status = "Folder Created";
            }
            else
            {
                status = "The was a problem when creating the folder No: " ;
            }
            return status;

            

            
        } //Create Folder Works

        internal object Execute(IRestRequest request)
        {
            throw new NotImplementedException();
        }

        //name and FileToBeUploaded can be same coz name is basically renaming the same File
        public string uploadDocument(string path,string name, string FileToBeUploaded,string parent_id ,int type, string ticket)
        {


            //string ticket = Authenticate();
            // Uses RestSharp as a import and calls it's RestClient method
            RestClient client = new RestClient(endPoint);


           // path = @"C:\Users\pndimande\Desktop\Travel Request TemplateTest.docx"; this was testing

            byte[] bytes;


            using (var fs = File.OpenRead(path))
            {
                using (var ms = new MemoryStream())
                {
                    ms.SetLength(fs.Length);
                    fs.CopyTo(ms);
                    bytes = ms.ToArray();
                }
            }


           
            var request = new RestRequest("/api/v2/nodes",Method.POST)
            {
               // string FileToBeUploaded = "Travel Request TemplateTest.docx"
                AlwaysMultipartFormData = true,
                Files = { FileParameter.Create("file", bytes, FileToBeUploaded) }
            };

            request.AddHeader("otcsticket", ticket);

            request.AddParameter("type", type);
            request.AddParameter("parent_id", parent_id);
            request.AddParameter("name", name);


            var response = client.Execute(request);

            // This will deserialize the received JSON in a Dictionary - key value pair
            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);

            string status = response.StatusCode.ToString();
            if (status == "OK")
            {
                status = "Document Uploaded";
            }
            else
            {
                status = "There was a problem uploading a file, server or that folder does not exist.";
            }
            return status;

        } //Upload Doc Works

        public string moveNode(string id ,string parent_id, string ticket)
        {
                //string ticket = Authenticate();
                var client = new RestClient(endPoint);
                var request = new RestRequest("/api/v1/nodes/{25350}", Method.PUT);
          
                request.AddHeader("otcsticket", ticket);
                request.AddParameter("id", id);
                request.AddParameter("parent_id", parent_id);
                IRestResponse response = client.Execute(request);
                string status = response.StatusCode.ToString();
                return status;
        } //Move Node Works

        public string search(string nodeID, string ticket)
        {
            string folder_id = string.Empty;
           // string ticket = Authenticate();
            // folderName = "Folderr22";

            RestClient client = new RestClient(endPoint);
            var request = new RestRequest("/api/v2/search", Method.POST);
            request.AddParameter("where", nodeID);
            request.AddParameter("limit", 1);

            request.AddHeader("otcsticket", ticket);

            // Executes the response that you created
            var response = client.Execute(request);

            // This will deserialize the received JSON in a Dictionary - key value pair
            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);
            string status = response.StatusCode.ToString();
            if (status == "OK")
            {
                status = status = "File Found" ;
            }

            //return status;
            return status;
        } //Search Works

        public string GetAncestors(string NodeID, string ticket)
        {
           
            // string ticket = Authenticate();
            // folderName = "Folderr22";

            RestClient client = new RestClient(endPoint);
            var request = new RestRequest("/api/v1/nodes/{id}/ancestors", Method.GET);
            request.AddParameter("id", NodeID);
            request.AddHeader("otcsticket", ticket);

            // Executes the response that you created
            //var response = client.Execute(request);
            var response = client.Execute(request).Content;

            //string[] array = responseHttp[0];

            // This will deserialize the received JSON in a Dictionary - key value pair
            //var deserialize = new JsonDeserializer();
            //var output = deserialize.Deserialize<Dictionary<string, string>>(response);
            //string status = response.StatusCode.ToString();
            //if (status == "OK")
            //{
            //    status = status = "File Found";
            //}

            int id = 0;
            RESTAPITest1.Poco item = JsonConvert.DeserializeObject<RESTAPITest1.Poco>(response);

        

            foreach (var obj in item.Results)
            {
                // Getting the id from the seach results and setting it to the variable idS

                if(obj.Data.properties.Type_name.Equals("Business Workspace"))
                {
                    id = obj.Data.properties.Id;
                }
               

                // Inserting the id value to be used for changing the category attribute
               // InsertAttributeValues(ticket, id, changeValue, txtResponse, catAttrID, ++i, endpoint);
            }
            return id.ToString();
        }//Works

        public string getCategoryIDAndCategoryAttributeID( string Catagory_ID, string ticket)
        {
           // string categoryIDAndCategoryAttributeID = string.Empty;

            RestClient client = new RestClient(endPoint);
            
            var request = new RestRequest("/api/v2/nodes/{id}/categories", Method.GET);
           
            request.AddHeader("otcsticket", ticket);
            request.AddParameter("id", Catagory_ID);


            // Executes the response that you created
            var response = client.Execute(request).Content;
            

            // This will deserialize the received JSON in a Dictionary - key value pair
          //  var deserialize = new JsonDeserializer();
          //  var output = deserialize.Deserialize<Dictionary<string, string>>(response);
           // string status = response.StatusCode.ToString();
          //  if (status == "OK")
           // {
              //  status = status = "Category Updated";
          //  }

            return response;

        }

        public string ReturnNodeIDSearch(string nodeID, string ticket)

        {
            var id = 0;
           //string ticket = Authenticate(username,password);

            // Setting the URL with for the search method with its endpoint
            var client = new RestClient(endPoint);

            var request = new RestRequest("/api/v2/search", Method.POST).
                AddParameter("where", nodeID);
           
            request.AddHeader("otcsticket", ticket);


            var response = client.Execute(request).Content;


            RESTAPITest1.Poco item = JsonConvert.DeserializeObject<RESTAPITest1.Poco>(response);

            // countLabel.Text =countLabel.Text = (item.Results.Count).ToString();

            
           
            foreach (var obj in item.Results)
            {
                // Getting the id from the seach results and setting it to the variable idS
                if(obj.Data.properties.Type_name==("Business Workspace"))
                {
                    id = obj.Data.properties.Id;
                }
                else if(obj.Data.properties.Type_name == "Folder" && obj.Data.properties.Type == 0)
                {

                    id = obj.Data.properties.Id;
                }
               

            }

            return id.ToString();


        }// EOM  This Search Returns a the node ID 


        
        //The following codes is for update catagory in bulk

        public void search( string attributeID, string changeValue, TextBox txtResponse,string catAttrID, string ticket, Label countLabel)
        {
            int id = 0;
            //string ticket = Authenticate(username,password);
          

            // Setting the URL with for the search method with its endpoint
            var client = new RestClient(endPoint);

            var request = new RestRequest("/api/v2/search", Method.POST).
            AddParameter("where", attributeID);
            request.AddHeader("otcsticket", ticket);


            var response = client.Execute(request).Content;

            RESTAPITest1.Poco item = JsonConvert.DeserializeObject<RESTAPITest1.Poco>(response);


            countLabel.Text = (item.Results.Count).ToString();

            int i = 0;

            foreach (var obj in item.Results)
            {
                // Getting the id from the seach results and setting it to the variable idS
                id = obj.Data.properties.Id;

                // Inserting the id value to be used for changing the category attribute
                InsertAttributeValues(ticket, id, changeValue, txtResponse, catAttrID, ++i, endPoint);
            }

        }// EOM

        /// <summary>
        /// This method will insert the values you want into a particular categories attribute
        /// </summary>
        public void InsertAttributeValues(string ticket, int id, string changeValue, TextBox txtResponse, string attributeID, int i, string endpoint)
        {


            // Setting the endpoint for this method
            var client = new RestClient(endPoint);
            client.BaseUrl = new Uri(string.Format("{0}/api/v2/nodes/{1}/categories/{2}", endpoint, id, splitAttributeID(attributeID)));

            // Creating a string for the data that should be sent as payload to the enpoint
            //string send = string.Format("{" + @"""2608_2""" + ":" + @"""{0}""" + "}", changeValue);
            string send = "{" + string.Format(@"""{0}""", attributeID) + ":" + string.Format(@"""{0}""", changeValue) + "}";

            // Creating a request as a PUT method and sending it through with a body parameter
            var request = new RestRequest(Method.PUT).
                AddParameter("body", send);
            request.AddHeader("otcsticket", ticket);

            // Executing the response
            var response = client.Execute(request).Content;

            txtResponse.Text += string.Format("{0}: ", i);
            txtResponse.Text += response;
            txtResponse.Text += System.Environment.NewLine;


        }// EOM
        public void insertValues(string ticket, string id, string changeValue,string changeValue2)
        {
            string attributeID = "26007_5";
            string attributeID2 = "2608_3";
            // Setting the endpoint for this method
            var client = new RestClient(endPoint);
           

            // Creating a string for the data that should be sent as payload to the enpoint
            //string send = string.Format("{" + @"""2608_2""" + ":" + @"""{0}""" + "}", changeValue);
            string send = "{" + string.Format(@"""{0}""", attributeID) + ":" + string.Format(@"""{0}""", changeValue) + 
                ","+ string.Format(@"""{0}""", attributeID2) + ":" + string.Format(@"""{0}""", changeValue2)+
                "}";

            // Creating a request as a PUT method and sending it through with a body parameter
            var request = new RestRequest("/api/v2/nodes/{id}/categories/{att}",Method.PUT).
                AddParameter("body", send)
                .AddParameter("id",id)
                .AddParameter("att", splitAttributeID(attributeID));

            request.AddHeader("otcsticket", ticket);

            // Executing the response
            var response = client.Execute(request).Content;

           // txtResponse.Text += string.Format("{0}: ", i);
           // txtResponse.Text += response;
           // txtResponse.Text += System.Environment.NewLine;
        }

        public void insertValuesForHR(string ticket, string id, string changeValue)
        {
            string attributeID = "26007_5";
            //string attributeID2 = "26007_6";
            // Setting the endpoint for this method
            var client = new RestClient(endPoint);


            // Creating a string for the data that should be sent as payload to the enpoint
            //string send = string.Format("{" + @"""2608_2""" + ":" + @"""{0}""" + "}", changeValue);
            string send = "{" + string.Format(@"""{0}""", attributeID) + ":" + string.Format(@"""{0}""", changeValue) +"}"; 
                
                //+
                //"," + string.Format(@"""{0}""", attributeID2) + ":" + string.Format(@"""{0}""", changeValue2) +
                //"}";

            // Creating a request as a PUT method and sending it through with a body parameter
            var request = new RestRequest("/api/v2/nodes/{id}/categories/{category_id}", Method.PUT).
             
                AddParameter("id", id)
                .AddParameter("category_id", splitAttributeID(attributeID)).
                 AddParameter("body", send);


            request.AddHeader("otcsticket", ticket);

            // Executing the response
            var response = client.Execute(request).Content;

            // txtResponse.Text += string.Format("{0}: ", i);
            // txtResponse.Text += response;
            // txtResponse.Text += System.Environment.NewLine;
        }



        public string UpdateFolderName(string ticket, string nodeID,string change)
        {

           

            // Setting the URL with for the search method with its endpoint
            var client = new RestClient(endPoint);

            var request = new RestRequest("/api/v2/nodes/{id}", Method.PUT).
            AddParameter("id", nodeID).
            AddParameter("name", change);
            request.AddHeader("otcsticket", ticket);


            var response = client.Execute(request);
            var deserialize = new JsonDeserializer();
            var output = deserialize.Deserialize<Dictionary<string, string>>(response);
            string status = response.StatusCode.ToString();
            if (status == "OK")
            {
                status = "Folder Updated";
            }
            else
            {
                status = "The was a problem when creating the folder No ";
            }

            return status;
        } //Please test this

        public string SubNodes(string ticket, string nodeID)
        {

            var id = 0;
            // Setting the URL with for the search method with its endpoint
            var client = new RestClient(endPoint);

            var request = new RestRequest("/api/v1/nodes/{id}/nodes", Method.GET).
            AddParameter("id", nodeID);
            request.AddHeader("otcsticket", ticket);


           var response =  client.Execute(request).Content;



            RESTAPITest1.Poco item = JsonConvert.DeserializeObject<RESTAPITest1.Poco>(response);

            // countLabel.Text =countLabel.Text = (item.Results.Count).ToString();



            foreach (var obj in item.Results)
            {
                // Getting the id from the seach results and setting it to the variable idS
                if (obj.Data.properties.Name == "10 - Appointments") 
                {
                    id = obj.Data.properties.Id; 
                }
               
 

            }
            return id.ToString() ;
        }


        /// <summary>
        /// This method will split the cat-attr-ID so that a cat-ID is returned
        /// </summary>
        public string splitAttributeID(string catAttrID)
        {
            string[] final = catAttrID.Split('_');
            return final[0];
        }

        public string Permissions(string dataID, string right_id,string ticket)
        {
            //string ticket = Authenticate();
            var client = new RestClient(endPoint);

         
            string permit = string.Format("see");
            string permit2 = string.Format("see_contents");

            string send = "{" + string.Format(@"""{0}""", "permissions") + ":" +"[" + string.Format(@"""{0}""", permit +"\"" +"," + "\"" + permit2 ) + "]"+ "}";

            //+
            //"," + string.Format(@"""{0}""", attributeID2) + ":" + string.Format(@"""{0}""", changeValue2) +
            //"}";

            // Creating a request as a PUT method and sending it through with a body parameter
            var request = new RestRequest("/api/v2/nodes/{id}/permissions/custom/{right_id}", Method.PUT).

                AddParameter("id", dataID)
                .AddParameter("right_id", right_id)
                .AddParameter("body", send);
        
            request.AddHeader("otcsticket", ticket);

            // Executing the response
            var response = client.Execute(request);

            string status = response.StatusCode.ToString();

            if (status == "OK")
            {
                status = "Permission changed";
            }
            else
            {
                status = "There was a problem when changing the permission";
            }
            return status;

        }

    }

}

