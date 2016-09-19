# Project-JSOM-List-Projects-Tasks
Lists the Projects and Tasks within the PWA site


## Setup 

1.	Create one or more projects that have a set of tasks that you would like to re-use
2.	Publish the project

### Using App

1.	Choose the project and click Show Tasks


### Prerequisites/Deployment
To use this code sample, you need the following:
* Project Server 2013 or Project Online (with subscription)
* Visual Studio 2013 or later 
* App for SharePoint project type
* Update the project site Url (Project Property) to match the site you are testing against.
* If you are not using a Developer site collection, you may need to enable [Side Loading] (https://blogs.msdn.microsoft.com/officeapps/2013/12/10/enable-app-sideloading-in-your-non-developer-site-collection/)



## How the sample affects your tenant data
This sample runs CSOM methods that read project and read/create task data. Tenant data will be affected.

## Additional resources
* [PS namespace (ps.js)] (https://msdn.microsoft.com/en-us/library/office/jj669820.aspx)
* [Client-side object model (CSOM) for Project 2013] (https://msdn.microsoft.com/en-us/library/office/jj163123.aspx)
* [SharePoint and Project Online SDK] (https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.
