using System;
using System.AddIn;
using System.ComponentModel;
using System.Globalization;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Windows.Forms;
using AirportHoursButton.SOAPICCS;
using RightNow.AddIns.AddInViews;


namespace AirportHoursButton
{
    public class WorkspaceRibbonAddIn : Panel, IWorkspaceRibbonButton
    {
        private bool InDesignMode;
        private IRecordContext RecordContext { get; set; }
        private IGlobalContext GlobalContext { get; set; }
        private IGenericObject AirportObject { get; set; }
        private DateTime dateOpens { get; set; }
        private DateTime dateCloses { get; set; }

        private RightNowSyncPortClient clientRN;

        public WorkspaceRibbonAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext globalContext)
        {
            if (!inDesignMode)
            {
                GlobalContext = globalContext;
                InDesignMode = inDesignMode;
                this.RecordContext = RecordContext;
                RecordContext.Saving += new CancelEventHandler(RecordContext_Saving);
            }
        }

        private void RecordContext_Saving(object sender, CancelEventArgs e)
        {
            try
            {
                Init();
                AirportObject = RecordContext.GetWorkspaceRecord("CO$Airports") as IGenericObject;
                RecordContext.RefreshWorkspace();
                if (AirportObject != null)
                {
                    int IdAirport = Convert.ToInt32(AirportObject.Id);
                    if (IdAirport != 0)
                    {
                        if (!Allday(IdAirport))
                        {
                            if (!boolGetHours(IdAirport))
                            {
                                MessageBox.Show("You must have at least one normal hour");
                                e.Cancel = true;
                            }
                            if (!CheckExtraordinary(IdAirport))
                            {
                                MessageBox.Show("Extraordinary hours are not under the normal hours.");
                                e.Cancel = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw ex;
            }
        }

        private bool Allday(int Airport)
        {
            bool all = true;
            ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
            APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
            clientInfoHeader.AppID = "Query Example";
            String queryString = "SELECT HoursOpen24 FROM  CO.Airports WHERE ID=" + Airport;
            clientRN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 10000, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
            foreach (CSVTable table in queryCSV.CSVTables)
            {
                String[] rowData = table.Rows;
                foreach (String data in rowData)
                {
                    all = Convert.ToInt32(data) == 0 ? false : true;
                }
            }
            return all;
        }
        private bool boolGetHours(int Airport)
        {
            bool hours = false;
            ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
            APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
            clientInfoHeader.AppID = "Query Example";
            String queryString = "SELECT  OpensZULUTime,ClosesZULUTime FROM CO.Airport_WorkingHours WHERE Type = 25 AND Airports =" + Airport + " Order by CreatedTime ASC";
            clientRN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 1, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
            int i = 0;
            foreach (CSVTable table in queryCSV.CSVTables)
            {
                String[] rowData = table.Rows;
                foreach (String data in rowData)
                {
                    Char delimiter = '|';
                    String[] substrings = data.Split(delimiter);
                    dateOpens = DateTime.Parse("2018/01/01 " + substrings[0].Trim());
                    dateCloses = DateTime.Parse("2018/01/01 " + substrings[1].Trim());
                    i++;
                }
            }
            return i > 0 ? true : false;
        }
        private bool CheckExtraordinary(int Airport)
        {
            bool extra = true;
            ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
            APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
            clientInfoHeader.AppID = "Query Example";
            String queryString = "SELECT OpensZULUTime,ClosesZULUTime FROM CO.Airport_WorkingHours WHERE Type = 1 AND Airports =" + Airport + " Order by CreatedTime ASC";
            clientRN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 1, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
            foreach (CSVTable table in queryCSV.CSVTables)
            {
                String[] rowData = table.Rows;
                foreach (String data in rowData)
                {
                    Char delimiter = '|';
                    String[] substrings = data.Split(delimiter);
                    DateTime dateOpensCompare = DateTime.Parse("2018/01/01 " + substrings[0].Trim());
                    DateTime dateClosesCompare = DateTime.Parse("2018/01/01 " + substrings[1].Trim());
                    double totalminutesOpen = (dateClosesCompare - dateOpens).TotalMinutes;
                    double totalminutesClose = (dateOpensCompare - dateCloses).TotalMinutes;
                    if (totalminutesOpen >= 0 || totalminutesClose <= 0)
                    {

                        extra = false;
                    }
                }
            }
            return extra;
        }
        public new void Click()
        {

        }
        public bool Init()
        {
            try
            {
                bool result = false;
                EndpointAddress endPointAddr = new EndpointAddress(GlobalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));
                // Minimum required
                BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;
                binding.ReceiveTimeout = new TimeSpan(0, 10, 0);
                binding.MaxReceivedMessageSize = 1048576; //1MB
                binding.SendTimeout = new TimeSpan(0, 10, 0);
                // Create client proxy class
                clientRN = new RightNowSyncPortClient(binding, endPointAddr);
                // Ask the client to not send the timestamp
                BindingElementCollection elements = clientRN.Endpoint.Binding.CreateBindingElements();
                elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                clientRN.Endpoint.Binding = new CustomBinding(elements);
                // Ask the Add-In framework the handle the session logic
                GlobalContext.PrepareConnectSession(clientRN.ChannelFactory);
                if (clientRN != null)
                {
                    result = true;
                }
                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

    }

    [AddIn("Airport Hours", Version = "1.0.0.0")]
    public class WorkspaceRibbonButtonFactory : IWorkspaceRibbonButtonFactory
    {
        private IGlobalContext GlobalContext;
        public IWorkspaceRibbonButton CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceRibbonAddIn(inDesignMode, RecordContext, GlobalContext);
        }
        public System.Drawing.Image Image32
        {
            get { return Properties.Resources.airport32; }
        }
        public System.Drawing.Image Image16
        {
            get { return Properties.Resources.airport16; }
        }
        public string Text
        {
            get { return "Validate Airport Hours"; }
        }
        public string Tooltip
        {
            get { return "Validate Airport Hours"; }
        }
        public bool Initialize(IGlobalContext GlobalContext)
        {
            this.GlobalContext = GlobalContext;
            return true;
        }
    }
}