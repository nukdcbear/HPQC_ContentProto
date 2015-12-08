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
            //ota.PlayWithFavorites();
            ota.CreateFavoriteFilter();
            //ota.GetRequirementTypes();
            //int reqID = ota.GetRequirementTypeID("SyRs");
            //Console.WriteLine("Requirement Type ID for SyRS is " + reqID);
            //ota.addBug2Project("Dummy Defect HP QC");
            //ota.addRequirement2Project("Another Dummy SyRS HP QC", "SyRS");
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
            Console.WriteLine("Connected to: " + almURL);

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
            conn.ReleaseConnection();
        }

        /* Get the requirement types */
        public void GetRequirementTypes2()
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

        public void GetRequirementTypes()
        {
            ReqFactory reqFact = conn.ReqFactory;
            List reqTypes = reqFact.GetRequirementTypes();

            foreach (ReqType reqType in reqTypes)
            {
                Console.WriteLine("Requirement type: " + reqType.Name);
                Console.WriteLine("Requirement type ID: " + reqType.ID);
            }
        }

        /* Get the requirment type ID based upon the requirement name */
        public int GetRequirementTypeID(String almReqTypeName)
        {
            ReqFactory reqFact = conn.ReqFactory;
            List reqTypes = reqFact.GetRequirementTypes();

            foreach (ReqType reqType in reqTypes)
            {

                if (String.Compare(reqType.Name.ToUpper(), almReqTypeName.ToUpper()) == 0)
                {
                    return reqType.ID;
                }
            }

            return 999;
        }

        /* Create a new requirement */
        public void addRequirement2Project(String reqName, String reqTypeName)
        {
            reqF = (ReqFactory)conn.ReqFactory;
            reqItem = (Req)reqF.AddItem(DBNull.Value);

            int reqTypeID = GetRequirementTypeID(reqTypeName);

            DateTime saveNow = DateTime.Now;

            //Populate the mandatory fields.
            reqItem.Name = reqName;
            // TODO Must query HP QC project to get the actual REQ_TYPE - TPR_TYPE_ID to use reqItem.TypeId
            // Obsolete reqItem.Type = "System";
            reqItem.TypeId = reqTypeID.ToString();
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

        /* Create new filter for Requirements by Type */
        public void FilterReqsByType(String reqTypeName)
        {
            reqF = (ReqFactory)conn.ReqFactory;
            TDFilter ReqFilter = reqF.Filter;

            int reqTypeID = GetRequirementTypeID(reqTypeName);

            ReqFilter["RQ_TYPE_ID"] = reqTypeID.ToString();

        }

        /* Playing with Favorites to see what we get */
        public void PlayWithFavorites()
        {
            FavoriteFactory favF = conn.GetCommonFavoriteFactory();
            //Favorite favItem = favF.AddItem(DBNull.Value);

            List favList = favF.NewList("");

            foreach (Favorite fav in favList)
            {
                Console.WriteLine("Name: " + fav.Name);
                Console.WriteLine("Module: " + fav.Module);
                Console.WriteLine("Public: " + fav.Public);
                Console.WriteLine("ParentID: " + fav.ParentId);
                Console.WriteLine("FilterData: " + fav.FilterData);
                Console.WriteLine("LayoutData: " + fav.LayoutData);
            }
        }

        /* Create a Favorite filter */
        public void CreateFavoriteFilter()
        {
            FavoriteFactory favF = conn.GetCommonFavoriteFactory();
            TDFilter favFilter = favF.Filter;
            favFilter["FAV_NAME"] = "\"All Defects\"";

            List favsFiltered = favF.NewList(favFilter.Text);
            Console.WriteLine("Number of Favorites Found: " + favsFiltered.Count);

            FavoriteFolderFactory favFF = conn.GetCommonFavoriteFolderFactory();
            List favFolders = favFF.NewList("");

            foreach (FavoriteFolder favFolder in favFolders)
            {
                if (favFolder.Public && String.Compare(favFolder.Module, "3") == 0)
                {
                    Favorite fav = favFolder.FavoriteFactory.AddItem(DBNull.Value);

                    fav.Name = "ALMPC Created Filter";
                    fav.FilterData = "<Filter Entity=\"requirement\" KeepHierarchical=\"true\"><Where /><Sort /><Grouping /></Filter>";
                    fav.LayoutData = "<?xml version=\"1.0\" encoding=\"utf - 16\"?><Layout><ViewLogicalName>TREE</ViewLogicalName><VisibleColumns><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_REQ_NAME</PhysicalFieldName><Width> 200 </Width></Column><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_REQ_ID</PhysicalFieldName><Width>100</Width></Column><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_REQ_STATUS</PhysicalFieldName><Width>120</Width></Column><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_TARGET_REL</PhysicalFieldName><Width>100</Width></Column><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_REQ_AUTHOR</PhysicalFieldName><Width>100</Width></Column><Column><EntityType>requirement</EntityType><PhysicalFieldName>RQ_TARGET_RCYC</PhysicalFieldName><Width>100</Width></Column></VisibleColumns><AdditionalData></AdditionalData></Layout>";
                    fav.ParentId = favFolder.ID;
                    fav["FAV_IS_PUBLIC"] = "Y";
                    fav["FAV_MODULE"] = 3; // Requirements module
                    fav.Post();
                    fav.Refresh();

                    Console.WriteLine("Name: " + fav.Name);
                    Console.WriteLine("Module: " + fav.Module);
                    Console.WriteLine("Public: " + fav.Public);
                    Console.WriteLine("ParentID: " + fav.ParentId);
                    Console.WriteLine("FilterData: " + fav.FilterData);

                }
            }
        }
    }
}
