using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using Microsoft.Toolkit.Win32.UI.Controls.Interop.WinRT;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace Panel_Tracker
{
    public partial class OAuth : Form
    {
        private const string FORGE_CLIENT_ID = "pt9GWttHTlNeRWpwVGtKS7GcM4PrMJdk";
        private const string FORGE_CLIENT_SECRET = "NmCPCkeXsZVspGtF";
        private const string FORGE_CALLBACK_URL = "http://localhost:3006";
        private const string FORGE_BASE_URL = "https://developer.api.autodesk.com";
        private const string FORGE_SCOPE = "data:read data:write data:create data:search bucket:create bucket:read bucket:update bucket:delete"; // assuming a full scope

        public bool exit = true;
        public string code { get; set; }
        public string bearer { get; set; }

        public OAuth()
        {
            InitializeComponent();

            //wb.Dock = DockStyle.Fill;
            //wb.NavigateError += new WebBrowserNavigateErrorEventHandler(wb_NavigateError);
            //Controls.Add(wb);

            // this is a basic code sample, quick & dirty way to get the Authentication string
            string authorizeURL = FORGE_BASE_URL + string.Format(
                "/authentication/v1/authorize?response_type=code&client_id={0}&redirect_uri={1}&scope={2}",
                FORGE_CLIENT_ID, FORGE_CALLBACK_URL, System.Net.WebUtility.UrlEncode(FORGE_SCOPE));

            // now let's open the Authorize page.
            //wb.Navigate(authorizeURL);

            wv.NavigationCompleted += new EventHandler<WebViewControlNavigationCompletedEventArgs>(wv_nav);
            wv.Navigate(authorizeURL);
        }

        private void wv_nav(object sender, WebViewControlNavigationCompletedEventArgs e)
        {
            if (e.Uri.ToString().Contains("code"))
            {
                var query = HttpUtility.ParseQueryString(e.Uri.Query);
                code = query["code"];

                string url = FORGE_BASE_URL + "/authentication/v1/gettoken";
                string postBody = string.Format(
                "client_id={0}&client_secret={1}&grant_type=authorization_code&code={2}&redirect_uri={3}",
                FORGE_CLIENT_ID, FORGE_CLIENT_SECRET, code, FORGE_CALLBACK_URL);
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("content-type", "application/x-www-form-urlencoded");
                headers.Add("client_id", FORGE_CLIENT_ID);
                headers.Add("client_secret", FORGE_CLIENT_SECRET);
                headers.Add("grant_type", "authorization_code");
                headers.Add("code", code);
                headers.Add("redirect_uri", FORGE_CALLBACK_URL);
                wv.NavigationCompleted += new EventHandler<WebViewControlNavigationCompletedEventArgs>(wv_nav2);
                wv.Navigate(new Uri(url), System.Net.Http.HttpMethod.Post, postBody, headers);
            }
        }

        private void wv_nav2(object sender, WebViewControlNavigationCompletedEventArgs e)
        {
            string html = wv.InvokeScript("eval", new string[] { "document.documentElement.outerHTML;" });
            string output = "{" + html.Split('{', '}')[1] + "}";
            dynamic responseBody = JObject.Parse(output);

            var values = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(output);
            bearer = "Bearer " + values["access_token"];

            this.Close();
            exit = true;
        }
    }
}
