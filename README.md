# Organization Directory

## Summary

Search People from Organization Directory and show live persona card on hover.

![directory](/samples/react-directory/assets/react-directory1.png)

![directory](/samples/react-directory/assets/react-directory2.png)

![directory](/samples/react-directory/assets/react-directory3.png)

![directory](/samples/react-directory/assets/react-directory4.png)

![directory](/samples/react-directory/assets/react-directory5.png)

![directory](/samples/react-directory/assets/react-directory6.png)

![directory](./assets/react-directory-withPaging.png)

![directory](/samples/react-directory/assets/react-directory-teams1.png)

![directory](/samples/react-directory/assets/react-directory-teams2.png)

![directory](/samples/react-directory/assets/react-directory-teams3.png)

## Compatibility

| :warning: Important          |
|:---------------------------|
| Every SPFx version is only compatible with specific version(s) of Node.js. In order to be able to build this sample, please ensure that the version of Node on your workstation matches one of the versions listed in this section. This sample will not work on a different version of Node.|
|Refer to <https://aka.ms/spfx-matrix> for more information on SPFx compatibility.   |

![SPFx 1.20.0](https://img.shields.io/badge/SPFx-1.20.0-green.svg)
![Node.js v18](https://img.shields.io/badge/Node.js-v18-green.svg)
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Teams Incompatible](https://img.shields.io/badge/Teams-Incompatible-lightgrey.svg)
![Local Workbench Incompatible](https://img.shields.io/badge/Local%20Workbench-Incompatible-red.svg "The solution requires access to your organization directory")
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)
![Compatible with Remote Containers](https://img.shields.io/badge/Remote%20Containers-Compatible-green.svg)

## Applies to

* [SharePoint Framework](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [Microsoft 365 tenant](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](https://aka.ms/m365/devprogram)

## Web Part Properties

|Property |Type|Required| comments|
|--------------------|----|--------|----------|
|Title | Text| No|Web Part Title|
|searchFirstName | boolean|No| Lastname or Firstname search query|
|Properties to search | text | No | By default **FirstName,LastName,WorkEmail,Department** are used for search. You can add custom properties separated by comma.|
|Properties to sent as clear text | text | No | By default if the search key has empty spaces, its replaced with **+** before sending it to the search query. The search properties mentioned here will be sent without the empty space replacement.|
|Results per page | number | Yes | Number of people result to be displayed per page. Max of **20** is allowed, default of **10** is set.|
|Filter Query| text | No | Additional search filter such as `Department:IT` or `BaseOfficeLocation:London`.|

## Contributors

* [João Mendes](https://github.com/joaojmendes)
* [Peter Paul Kirschner](https://github.com/petkir)
* [Sudharsan Kesavanarayanan](https://github.com/sudharsank)
* [Abderahman Moujahid](https://github.com/Abderahman88)
* [Sandeep P S](https://github.com/Sandeep-FED)

## Version history

|Version|Date|Comments|
|-------|----|--------|
|1.0.0|July 29, 2019|Initial release|
|1.0.1|July 19, 2020|Bugfix and mock-service for workbench (`LivePersonaCard` not supported in workbench)|
|2.0.0|Sep 18 2020|React hooks, paging, dynamic search props, result alignment using office ui fabric stack.|
|3.0.0|Oct 17 2020|Minor fixes and add the additional web part property.|
|3.0.1|March 4 2021|Bugfix 'Sort People by'|
|3.0.2|Oct 3 2022|Minor styling fixes and people container position|
|3.0.3|Oct 4 2022|Fix for LivePersonaCard|
|3.0.4|Jun 20 2023|Upgrade to SPFx 1.17.3|
|3.0.5|Oct 05 2024|Upgrade to SPFx 1.20.0|
|4.0.0|Apr 06 2025|Upgrade to Fluent UI v9 & other improvements|

## Minimal Path to Awesome

* Clone this repository (or [download this solution as a .ZIP file](https://pnp.github.io/download-partial/?url=https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-directory) then unzip it)
* From your command line, change your current directory to the directory containing this sample (`react-directory`, located under `samples`)
* in the command line run:
  * `npm install`
  * `gulp build`
  * `gulp bundle --ship`
  * `gulp package-solution --ship`
  * Add to AppCatalog and deploy

> This sample can also be opened with [VS Code Remote Development](https://code.visualstudio.com/docs/remote/remote-overview). Visit <https://aka.ms/spfx-devcontainer> for further instructions.

## Help

We do not support samples, but we this community is always willing to help, and we want to improve these samples. We use GitHub to track issues, which makes it easy for  community members to volunteer their time and help resolve issues.

If you're having issues building the solution, please run [spfx doctor](https://pnp.github.io/cli-microsoft365/cmd/spfx/spfx-doctor/) from within the solution folder to diagnose incompatibility issues with your environment.

You can try looking at [issues related to this sample](https://github.com/pnp/sp-dev-fx-webparts/issues?q=label%3A%22sample%3A%20react-directory") to see if anybody else is having the same issues.

You can also try looking at [discussions related to this sample](https://github.com/pnp/sp-dev-fx-webparts/discussions?discussions_q=react-directory) and see what the community is saying.

If you encounter any issues while using this sample, [create a new issue](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Abug-suspected%2Csample%3A%20react-directory&template=bug-report.yml&sample=react-directory&authors=@joaojmendes%20@petkir%20@sudharsank%20@Abderahman88&title=react-directory%20-%20).

For questions regarding this sample, [create a new question](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Aquestion%2Csample%3A%20react-directory&template=question.yml&sample=react-directory&authors=@joaojmendes%20@petkir%20@sudharsank%20@Abderahman88&title=react-directory%20-%20).

Finally, if you have an idea for improvement, [make a suggestion](https://github.com/pnp/sp-dev-fx-webparts/issues/new?assignees=&labels=Needs%3A+Triage+%3Amag%3A%2Ctype%3Aenhancement%2Csample%3A%20react-directory&template=question.yml&sample=react-directory&authors=@joaojmendes%20@petkir%20@sudharsank%20@Abderahman88&title=react-directory%20-%20).


## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

<img src="https://m365-visitor-stats.azurewebsites.net/sp-dev-fx-webparts/samples/react-directory" />
