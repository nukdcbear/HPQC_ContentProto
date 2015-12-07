using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDAPIOLELib;

namespace HPQC_ContentProto
{
    class Program
    {
        static void Main(string[] args)
        {
            OtaApiActions ota = new OtaApiActions();

            // Simple command line argument processing; [0] = Server URL, [1] = User ID, [2] = User password
            if (args.Length < 3)
            {
                Console.WriteLine("Please enter required arguments; Server URL, User ID and User password!");
                Environment.Exit(1);
            }

            String almServerURL = args[0];
            String almUser = args[1];
            String almUserPassword = args[2];

            try
            {
                ota.connectToProject(almServerURL, "POC", "ALM", almUser, almUserPassword);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not establish a connection to HP QC " + ex.Message);
            }

            ota.GetRequirementTypes2();
            //ota.addBug2Project("Dummy Defect HP QC");
            //ota.addRequirement2Project("Dummy SyRS HP QC");
            ota.disconnect();

            Environment.Exit(0);
        }
    }

    class OtaApiActions
    {
        private TDConnection conn;
        private ReqFactory reqF; // Do Not Use the Interface use direct Objects
        private Req reqItem;
        IBugFactory2 bugF;
        IBug bugItem;

        /* Create a connection to an ALM server           */
        /* Authenticate the user and log on to a project  */
        public void connectToProject(String almURL, String almDomain, String almProject, String almUser, String almUserPasswd)
        {
            conn = new TDConnection();
            conn.InitConnectionEx(almURL);

            //Authentication
            try
            {
                conn.Login(almUser, almUserPasswd);
                Console.WriteLine("Logged in as " + almUser);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                //Connect to a project
                conn.Connect(almDomain, almProject);
                Console.WriteLine("Connected to project: " + almProject);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /* Create a pre-authenticated connection to an ALM server  */
        public void authenticatedConnectToServer(String almURL, String almDomain, String almProject, String almUser, String almUserPasswd)
        {
            TDConnection otaConnection = new TDConnection();

            // configure connection to add Basic Auth Header at first request 
            otaConnection.SetBasicAuthHeaderMode(TDAPI_BASIC_AUTH_HEADER_MODES.HEADER_MODE_ONCE);

            // set user credentials
            otaConnection.SetServerCredentials(almUser, almUserPasswd);

            // init connect to ALM server       
            otaConnection.InitConnection(almURL);

            // Connect to a project
            otaConnection.Connect(almDomain, almProject);
        }
        public void disconnect()
        {
            conn.Disconnect();
            conn.Logout();
        }

        /* Get the requirement types */
        public void GetRequirementTypes()
        {
            Customization almCustomization = (Customization)conn.Customization;
            almCustomization.Load();

            // Get the RequirementType objects one after another
            CustomizationTypes custTypes = (CustomizationTypes)almCustomization.Types;

            // Entity type of 0 is requirement type
            List reqTypes = custTypes.GetEntityCustomizationTypes(0);

            Console.WriteLine("Number of Requirement Types: " + reqTypes.Count);

            foreach (CustomizationReqType item in reqTypes)
            {
                Console.WriteLine("Requirement type: " + item.Name);
                Console.WriteLine("Requirement type ID: " + item.ID);
            }

        }

        public void GetRequirementTypes2()
        {
            ReqFactory reqFact = conn.ReqFactory;
            List reqTypes = reqFact.GetRequirementTypes();

            foreach (ReqType reqType in reqTypes)
            {
                Console.WriteLine("Requirement type: " + reqType.Name);
                Console.WriteLine("Requirement type ID: " + reqType.ID);
            }
        }

        /* Create a new requirement */
        public void addRequirement2Project(String reqName)
        {
            reqF = (ReqFactory)conn.ReqFactory;
            reqItem = (Req)reqF.AddItem(DBNull.Value);

            DateTime saveNow = DateTime.Now;

            //Populate the mandatory fields.
            reqItem.Name = reqName;
            // TODO Must query HP QC project to get the actual REQ_TYPE - TPR_TYPE_ID to use reqItem.TypeId
            // Obsolete reqItem.Type = "System";
            reqItem.TypeId = "101";  // We know the System requirement TypeId is 101 in the POC-ALM project.
            reqItem.Author = "dbarringer";
            reqItem.Priority = "2-Medium";

            //Post the requirement
            reqItem.Post();
        }

        /* Create a new defect */
        public void addBug2Project(String bugName)
        {
            bugF = (IBugFactory2)conn.BugFactory;
            bugItem = (IBug)bugF.AddItem(DBNull.Value);

            DateTime saveNow = DateTime.Now;

            // Populate the mandatory fields.
            bugItem.Summary = bugName;
            bugItem.DetectedBy = "dbarringer";
            bugItem.Priority = "2-Medium";
            bugItem.Status = "New";
            bugItem["BG_DETECTION_DATE"] = saveNow;
            bugItem["BG_SEVERITY"] = "2-Medium";
            bugItem["BG_DESCRIPTION"] = "This is a description for " + bugName;

            //Post the bug
            bugItem.Post();
        }
    }
}
