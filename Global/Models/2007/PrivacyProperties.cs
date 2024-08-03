using System;
using System.Net.Http;
using System.Text;
using System.Reflection;
using Newtonsoft.Json;

namespace OpenXMLOffice.Global_2007
{

    /// <summary>
    /// This class deals with anonymized data collection status flags and code flow details.
    /// User can opt-out any part of data collection and help us to understand the product use by sharing the data they are willing to share.
    /// No data will be used in marketing, sales or personal identification now or ever by this package
    /// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
    /// </summary>
    public class PrivacyProperties
    {
        private class StatsPayload
        {
            public bool UsageCounter { get; internal set; }
            public bool HostProcess64Bit { get; internal set; }
            public bool Os64Bit { get; internal set; }
            public string DotnetVersion { get; internal set; }
            public bool IsInteractiveApp { get; internal set; }
            public string OsVersion { get; internal set; }
            public bool EnableGeoTagging { get; internal set; }
            public object PackageName { get; internal set; }
            public string PackageVersion { get; internal set; }
            public string GlobalVersion { get; internal set; }
            public bool PackageDetails { get; internal set; }
            public bool HardwareDetails { get; internal set; }
        }
        /// <summary>
        /// 
        /// </summary>
        protected bool isFileEdited = false;
        private bool shareHardwareDetails = true;
        /// <summary>
        /// Turn off hardware details collection.
        /// Data collection related to type of Operating System, hardware resource details along with language and timezone setting
        /// This helps us in optimizing code logic to run of minimal general available hardware pattern. 
        /// Language and timezone helps us in understanding the language support that we have to support in our documents and office components.
        /// This details also influenced by geoLocation setting to decide on document efforts in different languages
        /// </summary>
        public bool ShareOsHardwareDetails
        {
            get
            {
                return shareHardwareDetails;
            }
            set
            {
                shareHardwareDetails = value;
            }
        }
        private bool shareIpGeoLocation = true;
        /// <summary>
        /// Turn off IP geo location details collection
        /// On request first reach the IP is converted to geo location and discarded
        /// Geo location data is aggregated and stored in the system and no IP will be ever stored at any point of time.
        /// This is Data is collected and stored as country and city level. As Aggregated counter providing no other valuable information about any personal detail. 
        /// This data helps us in enabling language support and font,char code consideration that has to be focused on.
        /// </summary>
        public bool ShareIpGeoLocation
        {
            get
            {
                return shareIpGeoLocation;
            }
            set
            {
                shareIpGeoLocation = value;
            }
        }
        private bool sharePackageRelatedDetails = true;
        /// <summary>
        /// Share the current type of package and its version you are using.
        /// This enables us to understand the package usage "spreadsheet"/"presentation"/"word" and the version that's widely used
        /// This helps us in making effort towards maintaining the backwards compatibility of new improvements
        /// </summary>
        public bool SharePackageRelatedDetails
        {
            get
            {
                return sharePackageRelatedDetails;
            }
            set
            {
                sharePackageRelatedDetails = value;
            }
        }
        private bool shareComponentRelatedDetails = true;
        /// <summary>
        /// Share the internal components used from a package
        /// We never share any details/data supplied when using the package
        /// This are information related to use of OpenXML components like table, picture etc.
        /// This enables us to understand the component most used and divert more efforts in working and improving the components that are most used
        /// </summary>
        public bool ShareComponentRelatedDetails
        {
            get
            {
                return shareComponentRelatedDetails;
            }
            set
            {
                shareComponentRelatedDetails = value;
            }
        }
        private bool shareUsageCounterDetails = true;
        /// <summary>
        /// This increment counter call that is made without any other data. Just giving data another file is generated thought our package.
        /// This will help us motivated community contribution and my sponsors informed about the impact and use of package that is getting the support
        /// </summary>
        public bool ShareUsageCounterDetails
        {
            get
            {
                return shareUsageCounterDetails;
            }
            set
            {
                shareUsageCounterDetails = value;
            }
        }
        /// <summary>
        /// Should only used by internal class
        /// </summary>
        protected PrivacyProperties()
        {
            var test4 = Assembly.GetExecutingAssembly();
        }

        /// <summary>
        /// Save data stats on save of new file
        /// </summary>
        protected void SendAnonymousSaveStates(AssemblyName assemblyName)
        {
            ShareUsageCounterDetails = true;
            ShareOsHardwareDetails = true;
            ShareIpGeoLocation = true;
            SharePackageRelatedDetails = true;
            ShareComponentRelatedDetails = true;
            StatsPayload statsPayload = new StatsPayload();
            if (ShareUsageCounterDetails)
            {
                statsPayload.UsageCounter = true;
                statsPayload.PackageName = assemblyName.FullName;
            }
            if (ShareOsHardwareDetails)
            {
                statsPayload.HardwareDetails = true;
                statsPayload.OsVersion = Environment.OSVersion.ToString();
                statsPayload.Os64Bit = Environment.Is64BitOperatingSystem;
                statsPayload.HostProcess64Bit = Environment.Is64BitProcess;
                statsPayload.DotnetVersion = Environment.Version.ToString();
                statsPayload.IsInteractiveApp = Environment.UserInteractive;
            }
            if (ShareIpGeoLocation)
            {
                statsPayload.EnableGeoTagging = true;
            }
            if (SharePackageRelatedDetails)
            {
                statsPayload.PackageDetails = true;
                statsPayload.GlobalVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                statsPayload.PackageName = assemblyName.FullName;
                statsPayload.PackageVersion = assemblyName.Version.ToString();
            }
            if (ShareComponentRelatedDetails)
            {
                // TODO : Collect used components details but without invading into too much details
            }
            SendPostData(JsonConvert.SerializeObject(statsPayload));
        }

        private async void SendPostData(string jsonDataToSend)
        {
            // Return If All Status sharing are blocked
            if (!ShareComponentRelatedDetails && !ShareOsHardwareDetails &&
            !ShareIpGeoLocation && !SharePackageRelatedDetails &&
            !ShareUsageCounterDetails)
            {
                return;
            }
            string url = "https://draviavemal.com/openxml-office/stats";
            using (HttpClient client = new HttpClient())
            {
                HttpContent content = new StringContent(jsonDataToSend, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(url, content);
                if (response.IsSuccessStatusCode)
                {
                    await response.Content.ReadAsStringAsync();
                }
            }
        }
    }
}