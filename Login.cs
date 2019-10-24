using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using RestClientTest;
using RestSharp;


namespace RESTAPITest1
{
    public partial class Login : Form
    {

        public bool BtnClick;
        public Login()
        {
            InitializeComponent();
            

        }
        


        RestTest restClient = new RestTest();

        public string username()
        {

            //txtuserName.Text = "pndimande";
            txtuserName.Text = "PTNN"; //Engen UserName
           
            string username = txtuserName.Text;
            return username;
        }
        public string password()
        {
            // txtPassword.Text = "Password2!";
            txtPassword.Text = "Datacentrix@2018"; //Engen Password
            
            string passwords = txtPassword.Text;
            return passwords;
        }

        private void button1_Click(object sender, EventArgs e)
        {

           
            
            string response = string.Empty;         
            response = restClient.Authenticate(username(),password());
           
            if (response.Length >15)
            {
                     this.Hide();
                     Form1 form1 = new Form1();
                     form1.Show();

              }
            if(txtPassword.Text == " " || txtuserName.Text == " ")
            {
                MessageBox.Show("Incorrect username or password!");
            }
             

        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
        private void ViewPassword(object sender, EventArgs e)
        {
            // If button is clicked
            if (!BtnClick)
            {
                BtnClick = true;

                // This changes the password to Plain Text
                txtPassword.PasswordChar = '\0';
                txtPassword.Focus();
            }
        }

        private void ViewPassword2(object sender, EventArgs e)
        {

            BtnClick = false;
            txtPassword.PasswordChar = '*';
            txtPassword.Focus();
        }

        private void txtuserName_TextChanged(object sender, EventArgs e)
        {
        
        }
    }
}
