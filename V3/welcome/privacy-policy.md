---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Privacy Policy

Repository URL [https://github.com/DraviaVemal/OpenXML-Office](https://github.com/DraviaVemal/OpenXML-Office)

### Introduction

Your privacy is our top priority. Any data collected will solely be used to enhance our understanding of usage patterns and identify which areas require the most attention. We gather anonymous usage statistics to improve how our package is utilized. This policy outlines how we collect, use, and safeguard your data.

### Data Collection

* **System Information:** Details about your operating system, hardware, and software environment.
* **IP Address:** Your IP address is collected for geographic and regional analytics but is anonymized to prevent identification of individual users.
* **Service Usage:** Information about the features and components of our package that are used.

### Purpose of Data Collection

We do not collect personally identifiable information (PII). The data collected is aggregated and anonymized to ensure that individual users cannot be identified. Your IP address is anonymized to prevent any linkage to your personal identity.

### Opt-Out

You can choose not to participate in sharing of usage statistics by updating the `PrivacyProperties` flags.

```csharp
PrivacyProperties.ShareComponentRelatedDetails = false;
PrivacyProperties.ShareIpGeoLocation = false;
PrivacyProperties.ShareOsDetails= false;
PrivacyProperties.SharePackageRelatedDetails = false;
PrivacyProperties.ShareUsageCounterDetails = false;
```

### More Specific Information

* **ShareComponentRelatedDetails**\
  Share the internal components used from a package. \
  We never share any details/data supplied when using the package. \
  This are information related to use of OpenXML components like table, picture etc. \
  This enables us to understand the component most used and divert more efforts in working and improving the components that are most used. \
  I'm still working on figuring out right composition of data that's useful for me to develop most used component and give the maximum extend of privacy on what is shared.\
  Note: Nothing is shared as of now
* **ShareIpGeoLocation**\
  Turn off IP geo location details collection.\
  On request first reach the IP is converted to geo location and discarded. Geo location data is aggregated and stored in the system and no IP will be ever stored at any point of time.\
  This Data is collected and stored at country and city level.\
  As Aggregated counter providing no other valuable information about any personal detail.\
  This data helps us in enabling language support and font,char code consideration that has to be focused on.
* **ShareOsDetails**\
  Turn off Operating system details collection.\
  Data collection related to type of Operating System, .net framework, process type.\
  This helps us decide minimum version and backward compatibility decision.
* **SharePackageRelatedDetails**\
  Share the current type of package and its version you are using. \
  This enables us to understand the package usage "spreadsheet"/"presentation"/"word" and the version that's widely used.\
  This helps us in making effort towards maintaining the backwards compatibility of new improvements.
* **ShareUsageCounterDetails**\
  This increment counter call that is made without any other data. Just triggering counter another file is generated thought our package.\
  This will help us motivated community contribution and my sponsors, informed about the impact and use of package that is getting the support.

### Changes to This Policy

We may update this policy from time to time. Any changes will be communicated via this privacy details document and will be effective from next version of the released package.\
If you want to know the terms and info about each version's privacy during its release time please refer to nuget package page.

Contact Information

For any questions or concerns about this policy, please contact me at contact@draviavemal.com
