using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using System.IO;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;
using System.Net;
using Microsoft.Identity.Client;
using System.Security;

namespace Office365_AssignLicenses_Console
{
    class Program
    {

        class Office365_License_Mapping
        {
            public string ProductName { get; set; }
            public string StringID { get; set; }
            public string GUIDService { get; set; }
            //           public string PlansIncluded_forfutureuse { get; set; }

        }

        static List<Office365_License_Mapping> License_Mapping_List = new List<Office365_License_Mapping>
        {
            new Office365_License_Mapping(){ProductName="",StringID="",GUIDService=""},
            new Office365_License_Mapping(){ProductName="AUDIO CONFERENCING",StringID="MCOMEETADV",GUIDService="0c266dff-15dd-4b49-8397-2bb16070ed52"},
            new Office365_License_Mapping(){ProductName="AZURE ACTIVE DIRECTORY BASIC",StringID="AAD_BASIC",GUIDService="2b9c8e7c-319c-43a2-a2a0-48c5c6161de7"},
            new Office365_License_Mapping(){ProductName="AZURE ACTIVE DIRECTORY PREMIUM P1",StringID="AAD_PREMIUM",GUIDService="078d2b04-f1bd-4111-bbd4-b4b1b354cef4"},
            new Office365_License_Mapping(){ProductName="AZURE ACTIVE DIRECTORY PREMIUM P2",StringID="AAD_PREMIUM_P2",GUIDService="84a661c4-e949-4bd2-a560-ed7766fcaf2b"},
            new Office365_License_Mapping(){ProductName="AZURE INFORMATION PROTECTION PLAN 1",StringID="RIGHTSMANAGEMENT",GUIDService="c52ea49f-fe5d-4e95-93ba-1de91d380f89"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION",StringID="DYN365_ENTERPRISE_PLAN1",GUIDService="ea126fc5-a19e-42e2-a731-da9d437bffcf"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION",StringID="DYN365_ENTERPRISE_CUSTOMER_SERVICE",GUIDService="749742bf-0d37-4158-a120-33567104deeb"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION",StringID="DYN365_FINANCIALS_BUSINESS_SKU",GUIDService="cc13a803-544e-4464-b4e4-6d6169a138fa"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION",StringID="DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE",GUIDService="8edc2cf8-6438-4fa9-b6e3-aa1660c640cc"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 FOR SALES ENTERPRISE EDITION",StringID="DYN365_ENTERPRISE_SALES",GUIDService="1e1a282c-9c54-43a2-9310-98ef728faace"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION",StringID="DYN365_ENTERPRISE_TEAM_MEMBERS",GUIDService="8e7a3d30-d97d-43ab-837c-d7701cef83dc"},
            new Office365_License_Mapping(){ProductName="DYNAMICS 365 UNF OPS PLAN ENT EDITION",StringID="Dynamics_365_for_Operations",GUIDService="ccba3cfe-71ef-423a-bd87-b6df3dce59a9"},
            new Office365_License_Mapping(){ProductName="ENTERPRISE MOBILITY + SECURITY E3",StringID="EMS",GUIDService="efccb6f7-5641-4e0e-bd10-b4976e1bf68e"},
            new Office365_License_Mapping(){ProductName="ENTERPRISE MOBILITY + SECURITY E5",StringID="EMSPREMIUM",GUIDService="b05e124f-c7cc-45a0-a6aa-8cf78c946968"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE (PLAN 1)",StringID="EXCHANGESTANDARD",GUIDService="4b9405b0-7788-4568-add1-99614e613b69"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE (PLAN 2)",StringID="EXCHANGEENTERPRISE",GUIDService="19ec0d23-8335-4cbd-94ac-6050e30712fa"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE",StringID="EXCHANGEARCHIVE_ADDON",GUIDService="ee02fd1b-340e-4a4b-b355-4a514e4c8943"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER",StringID="EXCHANGEARCHIVE",GUIDService="90b5e015-709a-4b8b-b08e-3200f994494c"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE ESSENTIALS",StringID="EXCHANGEESSENTIALS",GUIDService="7fc0182e-d107-4556-8329-7caaa511197b"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE ESSENTIALS",StringID="EXCHANGE_S_ESSENTIALS",GUIDService="e8f81a67-bd96-4074-b108-cf193eb9433b"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE KIOSK",StringID="EXCHANGEDESKLESS",GUIDService="80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82"},
            new Office365_License_Mapping(){ProductName="EXCHANGE ONLINE POP",StringID="EXCHANGETELCO",GUIDService="cb0a98a8-11bc-494c-83d9-c1b1ac65327e"},
            new Office365_License_Mapping(){ProductName="INTUNE",StringID="INTUNE_A",GUIDService="061f9ace-7d42-4136-88ac-31dc755f143f"},
            new Office365_License_Mapping(){ProductName="MICROSOFT 365 BUSINESS",StringID="SPB",GUIDService="cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"},
            new Office365_License_Mapping(){ProductName="MICROSOFT 365 E3",StringID="SPE_E3",GUIDService="05e9a617-0261-4cee-bb44-138d3ef5d965"},
            new Office365_License_Mapping(){ProductName="MICROSOFT DYNAMICS CRM ONLINE BASIC",StringID="CRMPLAN2",GUIDService="906af65a-2970-46d5-9b58-4e9aa50f0657"},
            new Office365_License_Mapping(){ProductName="MICROSOFT DYNAMICS CRM ONLINE",StringID="CRMSTANDARD",GUIDService="d17b27af-3f49-4822-99f9-56a661538792"},
            new Office365_License_Mapping(){ProductName="MICROSOFT INTUNE A DIRECT",StringID="INTUNE_A",GUIDService="061f9ace-7d42-4136-88ac-31dc755f143f"},
            new Office365_License_Mapping(){ProductName="MS IMAGINE ACADEMY",StringID="IT_ACADEMY_AD",GUIDService="ba9a34de-4489-469d-879c-0f0f145321cd"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS",StringID="O365_BUSINESS",GUIDService="cdd28e44-67e3-425e-be4c-737fab2899d3"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS",StringID="SMB_BUSINESS",GUIDService="b214fe43-f5a3-4703-beeb-fa97188220fc"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS ESSENTIALS",StringID="O365_BUSINESS_ESSENTIALS",GUIDService="3b555118-da6a-4418-894f-7df1e2096870"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS ESSENTIALS",StringID="SMB_BUSINESS_ESSENTIALS",GUIDService="dab7782a-93b1-4074-8bb1-0e61318bea0b"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS PREMIUM",StringID="O365_BUSINESS_PREMIUM",GUIDService="f245ecc8-75af-4f8e-b61f-27d8114de5f3"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 BUSINESS PREMIUM",StringID="SMB_BUSINESS_PREMIUM",GUIDService="ac5cef5d-921b-4f97-9ef3-c99076e5470f"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E1",StringID="STANDARDPACK",GUIDService="18181a46-0d4e-45cd-891e-60aabd171b4e"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E2",StringID="STANDARDWOFFPACK",GUIDService="6634e0ce-1a9f-428c-a498-f84ec7b8aa2e"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E3",StringID="ENTERPRISEPACK",GUIDService="6fd2c87f-b296-42f0-b197-1e91e994b900"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E3 DEVELOPER",StringID="DEVELOPERPACK",GUIDService="189a915c-fe4f-4ffa-bde4-85b9628d07a0"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E4",StringID="ENTERPRISEWITHSCAL",GUIDService="1392051d-0cb9-4b7a-88d5-621fee5e8711"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E5",StringID="ENTERPRISEPREMIUM",GUIDService="c7df2760-2c81-4ef7-b578-5b5392b571df"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 ENTERPRISE E5 WITHOUT AUDIO CONFERENCING",StringID="ENTERPRISEPREMIUM_NOPSTNCONF",GUIDService="26d45bd9-adf1-46cd-a9e1-51e9a5524128"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 F1",StringID="DESKLESSPACK",GUIDService="4b585984-651b-448a-9e53-3b10f069cf7f"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 MIDSIZE BUSINESS",StringID="MIDSIZEPACK",GUIDService="04a7fb0d-32e0-4241-b4f5-3f7618cd1162"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 PROPLUS",StringID="OFFICESUBSCRIPTION",GUIDService="c2273bd0-dff7-4215-9ef5-2c7bcfb06425"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 SMALL BUSINESS",StringID="LITEPACK",GUIDService="bd09678e-b83c-4d3f-aaba-3dad4abd128b"},
            new Office365_License_Mapping(){ProductName="OFFICE 365 SMALL BUSINESS PREMIUM",StringID="LITEPACK_P2",GUIDService="fc14ec4a-4169-49a4-a51e-2c852931814b"},
            new Office365_License_Mapping(){ProductName="ONEDRIVE FOR BUSINESS (PLAN 1)",StringID="WACONEDRIVESTANDARD",GUIDService="e6778190-713e-4e4f-9119-8b8238de25df"},
            new Office365_License_Mapping(){ProductName="ONEDRIVE FOR BUSINESS (PLAN 2)",StringID="WACONEDRIVEENTERPRISE",GUIDService="ed01faf2-1d88-4947-ae91-45ca18703a96"},
            new Office365_License_Mapping(){ProductName="POWER BI FOR OFFICE 365 ADD-ON",StringID="POWER_BI_ADDON",GUIDService="45bc2c81-6072-436a-9b0b-3b12eefbc402"},
            new Office365_License_Mapping(){ProductName="POWER BI PRO",StringID="POWER_BI_PRO",GUIDService="f8a1db68-be16-40ed-86d5-cb42ce701560"},
            new Office365_License_Mapping(){ProductName="PROJECT FOR OFFICE 365",StringID="PROJECTCLIENT",GUIDService="a10d5e58-74da-4312-95c8-76be4e5b75a0"},
            new Office365_License_Mapping(){ProductName="PROJECT ONLINE ESSENTIALS",StringID="PROJECTESSENTIALS",GUIDService="776df282-9fc0-4862-99e2-70e561b9909e"},
            new Office365_License_Mapping(){ProductName="PROJECT ONLINE PREMIUM",StringID="PROJECTPREMIUM",GUIDService="09015f9f-377f-4538-bbb5-f75ceb09358a"},
            new Office365_License_Mapping(){ProductName="PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT",StringID="PROJECTONLINE_PLAN_1",GUIDService="2db84718-652c-47a7-860c-f10d8abbdae3"},
            new Office365_License_Mapping(){ProductName="PROJECT ONLINE PROFESSIONAL",StringID="PROJECTPROFESSIONAL",GUIDService="53818b1b-4a27-454b-8896-0dba576410e6"},
            new Office365_License_Mapping(){ProductName="PROJECT ONLINE WITH PROJECT FOR OFFICE 365",StringID="PROJECTONLINE_PLAN_2",GUIDService="f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c"},
            new Office365_License_Mapping(){ProductName="SHAREPOINT ONLINE (PLAN 1)",StringID="SHAREPOINTSTANDARD",GUIDService="1fc08a02-8b3d-43b9-831e-f76859e04e1a"},
            new Office365_License_Mapping(){ProductName="SHAREPOINT ONLINE (PLAN 2)",StringID="SHAREPOINTENTERPRISE",GUIDService="a9732ec9-17d9-494c-a51c-d6b45b384dcb"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS CLOUD PBX",StringID="MCOEV",GUIDService="e43b5b99-8dfb-405f-9987-dc307f34bcbd"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS ONLINE (PLAN 1)",StringID="MCOIMP",GUIDService="b8b749f8-a4ef-4887-9539-c95b1eaa5db7"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS ONLINE (PLAN 2)",StringID="MCOSTANDARD",GUIDService="d42c793f-6c78-4f43-92ca-e8f6a02b035f"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS PSTN CONFERENCING",StringID="MCOMEETADV",GUIDService="0c266dff-15dd-4b49-8397-2bb16070ed52"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING",StringID="MCOPSTN2",GUIDService="d3b4fe1f-9992-4930-8acb-ca6ec609365e"},
            new Office365_License_Mapping(){ProductName="SKYPE FOR BUSINESS PSTN DOMESTIC CALLING",StringID="MCOPSTN1",GUIDService="0dab259f-bf13-4952-b7f8-7db8f131b28d"},
            new Office365_License_Mapping(){ProductName="VISIO ONLINE PLAN 1",StringID="VISIOONLINE_PLAN1",GUIDService="4b244418-9658-4451-a2b8-b5e2b364e9bd"},
            new Office365_License_Mapping(){ProductName="VISIO Online Plan 2",StringID="VISIOCLIENT",GUIDService="c5928f49-12ba-48f7-ada3-0d743a3601d5"},
            new Office365_License_Mapping(){ProductName="WINDOWS 10 ENTERPRISE E3",StringID="WIN10_PRO_ENT_SUB",GUIDService="cb10e6cd-9da4-4992-867b-67546b1db821"},
            new Office365_License_Mapping(){ProductName="WINDOWS 10 ENTERPRISE E3 Trial",StringID="Win10_VDA_E3",GUIDService="NULL"}
        };



        class Office365_Plan_Mapping
        {
            public string PlanName { get; set; }
            public string LicenceName { get; set; }
            public string GUIDService { get; set; }
            //           public string PlansIncluded_forfutureuse { get; set; }

        }


        static List<Office365_Plan_Mapping> Plan_Mapping_List = new List<Office365_Plan_Mapping>
        {
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="BPOS_S_TODO_2",GUIDService="c87f142c-d1e9-4363-8630-aaea9c4d9ae5"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="Deskless",GUIDService="8c7d2df8-86f0-4902-b2ed-a0458298f3b3"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="EXCHANGE_S_ENTERPRISE",GUIDService="efb87545-963c-4e0d-99df-69c6916d9eb0"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="FLOW_O365_P2",GUIDService="76846ad7-7776-4c40-a281-a386362dd1b9"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="FORMS_PLAN_E3",GUIDService="2789c901-c14e-48ab-a76a-be334d9d793a"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="MCOSTANDARD",GUIDService="0feaeb32-d00e-4d66-bd5a-43b5b83db82c"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="OFFICESUBSCRIPTION",GUIDService="43de0ff5-c92c-492b-9116-175376d08c38"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="POWERAPPS_O365_P2",GUIDService="c68f8d98-5534-41c8-bf36-22fa496fa792"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="PROJECTWORKMANAGEMENT",GUIDService="b737dad2-2f6c-4c65-90e3-ca563267e8b9"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="RMS_S_ENTERPRISE",GUIDService="bea4c11e-220a-4e6d-8eb8-8ea15d019f90"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="SHAREPOINTENTERPRISE",GUIDService="5dbe027f-2339-4123-9542-606e4d348a72"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="SHAREPOINTWAC",GUIDService="e95bec33-7c88-4a70-8e19-b10bd9d0c014"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="STREAM_O365_E3",GUIDService="9e700747-8b1d-45e5-ab8d-ef187ceec156"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="SWAY",GUIDService="a23b959c-7ce8-4e57-9140-b90eb88a9e97"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="TEAMS1",GUIDService="57ff2da0-773e-42df-b2af-ffb7a2317929"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E3",PlanName="YAMMER_ENTERPRISE",GUIDService="7547a3fe-08ee-4ccb-b430-5077c5041653"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="ADALLOM_S_O365",GUIDService="8c098270-9dd4-4350-9b30-ba4703f3b36b"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="BI_AZURE_P2",GUIDService="70d33638-9c74-4d01-bfd3-562de28bd4ba"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="BPOS_S_TODO_3",GUIDService="3fb82609-8c27-4f7b-bd51-30634711ee67"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="Deskless",GUIDService="8c7d2df8-86f0-4902-b2ed-a0458298f3b3"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="EQUIVIO_ANALYTICS",GUIDService="4de31727-a228-4ec3-a5bf-8e45b5ca48cc"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="EXCHANGE_ANALYTICS",GUIDService="34c0d7a0-a70f-4668-9238-47f9fc208882"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="EXCHANGE_S_ENTERPRISE",GUIDService="efb87545-963c-4e0d-99df-69c6916d9eb0"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="FLOW_O365_P3",GUIDService="07699545-9485-468e-95b6-2fca3738be01"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="FORMS_PLAN_E5",GUIDService="e212cbc7-0961-4c40-9825-01117710dcb1"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="LOCKBOX_ENTERPRISE",GUIDService="9f431833-0334-42de-a7dc-70aa40db46db"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="MCOEV",GUIDService="4828c8ec-dc2e-4779-b502-87ac9ce28ab7"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="MCOMEETADV",GUIDService="3e26ee1f-8a5f-4d52-aee2-b81ce45c8f40"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="MCOSTANDARD",GUIDService="0feaeb32-d00e-4d66-bd5a-43b5b83db82c"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="OFFICESUBSCRIPTION",GUIDService="43de0ff5-c92c-492b-9116-175376d08c38"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="POWERAPPS_O365_P3",GUIDService="9c0dab89-a30c-4117-86e7-97bda240acd2"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="PROJECTWORKMANAGEMENT",GUIDService="b737dad2-2f6c-4c65-90e3-ca563267e8b9"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="RMS_S_ENTERPRISE",GUIDService="bea4c11e-220a-4e6d-8eb8-8ea15d019f90"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="SHAREPOINTENTERPRISE",GUIDService="5dbe027f-2339-4123-9542-606e4d348a72"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="SHAREPOINTWAC",GUIDService="e95bec33-7c88-4a70-8e19-b10bd9d0c014"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="STREAM_O365_E5",GUIDService="6c6042f5-6f01-4d67-b8c1-eb99d36eed3e"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="SWAY",GUIDService="a23b959c-7ce8-4e57-9140-b90eb88a9e97"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="TEAMS1",GUIDService="57ff2da0-773e-42df-b2af-ffb7a2317929"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="THREAT_INTELLIGENCE",GUIDService="8e0c0a52-6a6c-4d40-8370-dd62790dcd70"},
            new Office365_Plan_Mapping(){LicenceName="OFFICE 365 ENTERPRISE E5",PlanName="YAMMER_ENTERPRISE",GUIDService="7547a3fe-08ee-4ccb-b430-5077c5041653"},   
        };


        public static GraphServiceClient client;

        class Office365_Licenze
        {
            public Guid? SKUID { get; set; }
            public string SKUNAME { get; set; }
            public Guid? ServicePlanId { get; set; }
            public string ServicePlanName { get; set; }

        }

        class Office365_Licenze_Users
        {
            public string Status { get; set; }
            public Guid? ServicePlanId { get; set; }
            public string Service { get; set; }
            public string License { get; set; }
            public string SKUNAME { get; set; }

        }


        class Office365_Users_class
        {
            public string ObjectID { get; set; }
            public string UPN { get; set; }
            public List<Office365_Licenze_Users> Licenze { get; set; }

            public Office365_Users_class()
            {
                Licenze = new List<Office365_Licenze_Users>();
            }

        }

        class Office365_Users_class_simplified
        {
            public string ObjectID { get; set; }
            public string UPN { get; set; }
            public string License_Type { get; set; }


        }


        class License_Profiles
        {
            public string ProfileName { get; set; }
            public string LicenceType { get; set; }
            public string enabledservices { get; set; }

        }

        class Users_Profiles
        {
            public string UserPrincipalName { get; set; }
            public string Profile { get; set; }
        }


        class Actions_Class
        {
            public string UserPrincipalName { get; set; }
            public string Action { get; set; }
            public string Profile { get; set; }

        }
        class Config
        {
            public bool ExportSimpleReportAtEveryRun { get; set; }
             // public bool Enforce { get; set; }
             // Setting not respected, will be used in future versions

            public bool LogOnlyMode { get; set; }
            public bool ReportDiffences { get; set; }
            public bool ExportActions { get; set; }
            public string ClientIdForUserAuthn { get; set; }

            public string Tenant { get; set; }
            public string username { get; set; }
            public string password { get; set; }
            public string DefaultUsageLocation { get; set; }

        }


        static List<Actions_Class> Actions_to_do = new List<Actions_Class>();

        static List<Office365_Licenze> Lista_Licenze = new List<Office365_Licenze>();
        static List<Office365_Users_class> Office365_Users = new List<Office365_Users_class>();
        static List<Office365_Users_class_simplified> Office365_Users_Simplified = new List<Office365_Users_class_simplified>();
        static List<Users_Profiles> UsersProfiles;
        static List<License_Profiles> ProfileConfig;
        static Config Configuration;


        static void Get_Licenses_SKU(GraphServiceClient client)
        {
            IGraphServiceSubscribedSkusCollectionPage Office365_SKUs = client.SubscribedSkus.Request().GetAsync().Result;
            foreach (SubscribedSku SKU in Office365_SKUs)
            {
                foreach (ServicePlanInfo servicePlanInfo in SKU.ServicePlans)
                {
                    var temp = new Office365_Licenze();
                    //      Console.WriteLine(SKU.SkuId + " " + SKU.SkuPartNumber);
                    //      Console.WriteLine(servicePlanInfo.ServicePlanId + " " + servicePlanInfo.ServicePlanName);

                    temp.ServicePlanId = servicePlanInfo.ServicePlanId;
                    temp.ServicePlanName = servicePlanInfo.ServicePlanName;
                    temp.SKUID = SKU.SkuId;
                    temp.SKUNAME = SKU.SkuPartNumber;
                    Lista_Licenze.Add(temp);
                }
            }
        }


        static void Get_Users_Status(GraphServiceClient client)
        {
            IGraphServiceUsersCollectionPage Office365_users = client.Users.Request().GetAsync().Result;
            foreach (User user in Office365_users)
            {
                //   Console.WriteLine(user.UserPrincipalName);
                Office365_Users_class temp = new Office365_Users_class();
                temp.UPN = user.UserPrincipalName;
                temp.ObjectID = user.Id;
        

                foreach (AssignedLicense lic in user.AssignedLicenses)
                {
        

                    if (user.AssignedPlans.Count() != 0)
                    {
                        foreach (AssignedPlan license in user.AssignedPlans)
                        {
                            if (license.CapabilityStatus != "Deleted") { 


                                Office365_Licenze_Users temp2 = new Office365_Licenze_Users();
                                temp2.Service = license.Service;
                      //          temp2.SKUNAME = Plan_Mapping_List.Find(x => x.GUIDService == license.ServicePlanId.ToString()).LicenceName;
                                temp2.SKUNAME = Lista_Licenze.Find(x => x.ServicePlanId == license.ServicePlanId).SKUNAME;
                       
                                try
                                {
                                    
                                    temp2.License = License_Mapping_List.Find(x => x.GUIDService == lic.SkuId.ToString()).ProductName;
                                }
                                catch { temp2.License = "MAPPING NOT AVAILABLE"; }

                                temp2.ServicePlanId = license.ServicePlanId;
                                temp2.Status = license.CapabilityStatus;
                                temp.Licenze.Add(temp2);
                                    // Console.WriteLine(license.CapabilityStatus + "  "+ license.Service+ "  " + license.ServicePlanId);
                            }
                        }
                    }

                }

                Office365_Users.Add(temp);
            }
        }


        static void Generate_End_User_Report()
        {
            foreach (Office365_Users_class user in Office365_Users)
            {
                var licenze = user.Licenze.Select(x => x.License).Distinct();
                foreach (var lic in licenze)
                {
                    var temp = new Office365_Users_class_simplified();
                    temp.ObjectID = user.ObjectID;
                    temp.UPN = user.UPN;
                    temp.License_Type = lic;
                    Office365_Users_Simplified.Add(temp);

                }

                if (licenze.Count() == 0) {
                    var temp = new Office365_Users_class_simplified();
                    temp.ObjectID = user.ObjectID;
                    temp.UPN = user.UPN;
                    temp.License_Type = "NULL";
                    Office365_Users_Simplified.Add(temp);

                }
            }

            if (Configuration.ExportSimpleReportAtEveryRun)
            {
                var file = "Output\\Simple_Report.csv";
                using (StreamWriter sw = new StreamWriter(file))
                {
                    string header = "ObjectID,UserPrincipalName,LicenseAssigned";
                    sw.WriteLine(header);
                    // iterates over the users
                    foreach (Office365_Users_class_simplified u in Office365_Users_Simplified)
                    {
                        // creates an array of the user's values
                        string[] values = { u.ObjectID, u.UPN, u.License_Type };
                        // creates a new line
                        string line = String.Join(",", values);
                        // writes the line
                        sw.WriteLine(line);
                    }
                    // flushes the buffer
                    sw.Flush();
                }
                Trace.WriteLine("Simple Report Exported correctly in the Output folder");
            }
        }


        static void Read_Config_Files()
        {


            string curFile = @"Config\\profilesdefinition.json";

            if (!System.IO.File.Exists(curFile))
            {
                Trace.WriteLine("File profilesdefinition.json do not exist in folder Config");
                Console.WriteLine("File profilesdefinition.json do not exist in folder Config");
                System.Environment.Exit(-1);
            }



            curFile = @"Config\\usersdefinition.json";

            if (!System.IO.File.Exists(curFile))
            {
                Trace.WriteLine("File usersdefinition.json do not exist in folder Config");
                Console.WriteLine("File usersdefinition.json do not exist in folder Config");
                System.Environment.Exit(-1);
            }


            curFile = @"Config\\config.json";

            if (!System.IO.File.Exists(curFile))
            {
                Trace.WriteLine("File config.json do not exist in folder Config");
                Console.WriteLine("File config.json do not exist in folder Config");
                System.Environment.Exit(-1);
            }



            Trace.WriteLine("Reading Profiles JSON");
            using (StreamReader r = new StreamReader("Config\\profilesdefinition.json"))
            {
                string json = r.ReadToEnd();
                ProfileConfig = JsonConvert.DeserializeObject<List<License_Profiles>>(json);
            }


            Trace.WriteLine("Reading User Definition JSON");

            using (StreamReader r = new StreamReader("Config\\usersdefinition.json"))
            {
                string json = r.ReadToEnd();
                UsersProfiles = JsonConvert.DeserializeObject<List<Users_Profiles>>(json);
            }

            Trace.WriteLine("Reading General Configuration JSON");
            using (StreamReader r = new StreamReader("Config\\config.json"))
            {
                string json = r.ReadToEnd();
                Configuration = JsonConvert.DeserializeObject<Config>(json);
            }


        }

        static void Evaluate_Changes()
        {
            foreach (Users_Profiles user in UsersProfiles)
            {

                var find = Office365_Users_Simplified.Find(x => x.UPN == user.UserPrincipalName);
                var ConfiguredProfile = (ProfileConfig.Find(x => x.ProfileName == user.Profile)).LicenceType;
                if (find != null)
                {
                    if (find.License_Type != ConfiguredProfile)
                    {
                        Trace.WriteLine("Profile for " + user.UserPrincipalName + " is different....adding action to add license, I'll remove all the others");

                        Actions_Class temp = new Actions_Class();
                        temp.UserPrincipalName = user.UserPrincipalName;
                        temp.Action = "REMOVE";
                        temp.Profile = find.License_Type;
                        Actions_to_do.Add(temp);

                        Actions_Class temp2 = new Actions_Class();
                        temp2.UserPrincipalName = user.UserPrincipalName;
                        temp2.Action = "ADD";
                        temp2.Profile = user.Profile;
                        Actions_to_do.Add(temp2);
                    }
                    else
                    {
                        Trace.Write("Profile for " + user.UserPrincipalName + " match....verifying services....");

                        var find2 = Office365_Users.Find(x => x.UPN == user.UserPrincipalName);
                        var UserProfile = ProfileConfig.Find(x => x.ProfileName == user.Profile);


                        List<string> ConfiguredServicesList = new List<string>();
                        foreach (var service in find2.Licenze)
                        {
                            if (service.Status != "Deleted") { 
                                ConfiguredServicesList.Add(service.Service.ToLower());
                            }
                        }
                        ConfiguredServicesList = ConfiguredServicesList.Distinct().ToList();
                        ConfiguredServicesList = ConfiguredServicesList.OrderBy(x => x).ToList();


                        List<string> ExpectedServicesList = new List<string>();
                        foreach (var service in UserProfile.enabledservices.Split(new string[] { ";" }, StringSplitOptions.None))
                        {
                            ExpectedServicesList.Add(service.ToLower());
                        }
                        ExpectedServicesList = ExpectedServicesList.Distinct().ToList();
                        ExpectedServicesList = ExpectedServicesList.OrderBy(x => x).ToList();


                        if (ConfiguredServicesList.Count() != ExpectedServicesList.Count())
                        {
                            Trace.WriteLine("Services do not match...removing and readding");
                            Actions_Class temp = new Actions_Class();
                            temp.UserPrincipalName = user.UserPrincipalName;
                            temp.Action = "REMOVE";
                            temp.Profile = find.License_Type;
                            Actions_to_do.Add(temp);

                            Actions_Class temp2 = new Actions_Class();
                            temp2.UserPrincipalName = user.UserPrincipalName;
                            temp2.Action = "ADD";
                            temp2.Profile = user.Profile;
                            Actions_to_do.Add(temp2);
                        }
                        else
                        {
                            var firstNotSecond = ConfiguredServicesList.Except(ExpectedServicesList, StringComparer.OrdinalIgnoreCase).ToList();
                            var secondNotFirst = ExpectedServicesList.Except(ConfiguredServicesList).ToList();
                            if (firstNotSecond.Any() && secondNotFirst.Any())
                            {
                                Trace.WriteLine("Services do not match...removing and readding");
                                Actions_Class temp = new Actions_Class();
                                temp.UserPrincipalName = user.UserPrincipalName;
                                temp.Action = "REMOVE";
                                temp.Profile = find.License_Type;
                                Actions_to_do.Add(temp);

                                Actions_Class temp2 = new Actions_Class();
                                temp2.UserPrincipalName = user.UserPrincipalName;
                                temp2.Action = "ADD";
                                temp2.Profile = user.Profile;
                                Actions_to_do.Add(temp2);
                            } else
                            {
                                Trace.WriteLine("Services match...skipping");
                            }
                        }



                    }
                }
                else
                {
                    Trace.WriteLine(user.UserPrincipalName + " not found as user in the tenant");
                }

            }

            Trace.WriteLine("Changes evaluation finished");



            if (Configuration.ExportActions)
            {
                var file = "Output\\Action_Export.csv";
                using (StreamWriter sw = new StreamWriter(file))
                {
                    string header = "UserPrincipalName,Action,Profile";
                    sw.WriteLine(header);
                    // iterates over the users
                    foreach (var u in Actions_to_do)
                    {
                        // creates an array of the user's values
                        string[] values = { u.UserPrincipalName, u.Action, u.Profile };
                        // creates a new line
                        string line = String.Join(",", values);
                        // writes the line
                        sw.WriteLine(line);
                    }
                    // flushes the buffer
                    sw.Flush();
                }
                Trace.WriteLine("Action List Exported correctly in the Output folder");
            }


        }



        static async Task<string> Add_LicenseAsync(Actions_Class action, string token)
        {

            var UsageLocationUrl = "https://graph.microsoft.com/beta/users/" + action.UserPrincipalName;
            string ULjson = "{\"usageLocation\":\"" + Configuration.DefaultUsageLocation + "\" }";


            var ULresponse = await Patch_RequestAsync(UsageLocationUrl, ULjson, token);
            

            var ConfiguredProfile = (ProfileConfig.Find(x => x.ProfileName == action.Profile)).LicenceType;
            var skuid = License_Mapping_List.Find(x => x.ProductName == ConfiguredProfile).GUIDService;

            var url = "https://graph.microsoft.com/beta/users/" + action.UserPrincipalName + "/assignLicense";


            var enabled_services = (ProfileConfig.Find(x => x.ProfileName == action.Profile)).enabledservices.Split(new string[] { ";" }, StringSplitOptions.None);
            var search_plans_to_disable = Plan_Mapping_List.FindAll(x => x.LicenceName == ConfiguredProfile);


            foreach (var s in enabled_services)
            {
                search_plans_to_disable.RemoveAll(x => x.PlanName.Contains(s.ToUpper()));

            }

            string create_disabledPlans_string="";
            foreach (var s in search_plans_to_disable)
            {
                create_disabledPlans_string += "\""+s.GUIDService+"\",";

            }

            create_disabledPlans_string=create_disabledPlans_string.Remove(create_disabledPlans_string.Length - 1);

         
            try
            {
                string json = "";
                if (enabled_services[0] == "ALL")
                {
                     json = "{\"addLicenses\": [{\"disabledPlans\": [], \"skuId\": \"" + skuid + "\"}],\"removeLicenses\": []}";
                }
                else { 
                     json = "{\"addLicenses\": [{\"disabledPlans\": ["+ create_disabledPlans_string + "], \"skuId\": \"" + skuid + "\"}],\"removeLicenses\": []}";
                }
                var response = await Post_RequestAsync(url, json, token);
                return response;
            }
            catch (Exception e)
            {
                Trace.WriteLine("Cannot add licence " + action.Profile + " for the user " + action.UserPrincipalName + " ...sorry....something wrong!");
                Trace.WriteLine(e.Message);
                Trace.WriteLine(e.InnerException);
                Trace.WriteLine(e.StackTrace);
                return "KO";
            }
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        static async Task<string> Post_RequestAsync(string url, string json, string token)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            var client = new HttpClient();

            //setup client
            client.BaseAddress = new Uri(url);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            var theContent = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
           
            var response = client.PostAsync(url, theContent).Result;
            
            var responseString = response.Content.ReadAsStringAsync();
            return responseString.Result;

        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        static async Task<string> Patch_RequestAsync(string url, string json, string token)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {

            
            var client = new HttpClient();

            //setup client
            client.BaseAddress = new Uri(url);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

            var method = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(method, client.BaseAddress)
            {
                Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json")
            };


            var response = client.SendAsync(request).Result;

            var responseString = response.Content.ReadAsStringAsync();
            return responseString.Result;
        }


        static async Task<string> Remove_LicenseAsync(Actions_Class action, string token)
        {
            var skuid = License_Mapping_List.Find(x => x.ProductName == action.Profile).GUIDService;
            var url = "https://graph.microsoft.com/beta/users/" + action.UserPrincipalName + "/assignLicense";

            try
            {
                if (skuid != "NULL")
                {
                    string json = "{addLicenses: [],removeLicenses: [\"" + skuid + "\"]}";
                    var response = await Post_RequestAsync(url, json, token);
                    return response;

                }
                else
                {
                    Trace.WriteLine("Cannot remove licence " + action.Profile + " for the user " + action.UserPrincipalName + " because GUID Service was not found .... GUID Mapping upgrade is required");
                    return "KO";
                }
            }
            catch (Exception e)
            {
                Trace.WriteLine("Cannot remove licence " + action.Profile + " for the user " + action.UserPrincipalName + " ...sorry....something wrong!");
                Trace.WriteLine(e.Message);
                Trace.WriteLine(e.InnerException);
                Trace.WriteLine(e.StackTrace);
                return "KO";
            }

        }








        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.

        static GraphServiceClient GetAuthenticatedClientForUser(string username, string password)

        {
            // Create Microsoft Graph client.
            try
            {

                graphClient = new GraphServiceClient(
                    "https://graph.microsoft.com/beta/",
                    new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            var token = await GetTokenForUserAsync(username, password);
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                            // This header has been added to identify our sample in the Microsoft Graph service.  If extracting this code for your project please remove.
                            //          requestMessage.Headers.Add("redirecturi", redirectUri);

                        }));
                return graphClient;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Could not create a graph client: " + ex.Message);
            }
            return graphClient;
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        static async Task<string> GetTokenForUserAsync(string username, string password)
        {
            AuthenticationResult authResult;
            string AuthorityUri = "https://login.microsoftonline.com/" + Configuration.Tenant + "/oauth2/v2.0/token";
            PublicClientApplication IdentityClientApp = new PublicClientApplication(Configuration.ClientIdForUserAuthn, AuthorityUri);

            string[] Scopes = { "User.Read",
                                           "User.ReadWrite","User.ReadWrite.All",
                                           "User.ReadBasic.All",
                                           "Directory.Read.All",
                                           "Directory.ReadWrite.All",
                                            // Group.Read.All is an admin-only scope. It allows you to read Group details.
                                            // Uncomment this scope if you want to run the application with an admin account
                                            // and perform the group operations in the UserMode class.
                                            // You'll also need to uncomment the UserMode.UserModeRequests.GetDetailsForGroups() method.

                                            //"Group.Read.All" 

                                        };
            try
            {
                var securePassword = new SecureString();
                foreach (char c in password)
                    securePassword.AppendChar(c);

                // authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, IdentityClientApp.Users.First());
                authResult = await IdentityClientApp.AcquireTokenByUsernamePasswordAsync(Scopes, username, securePassword);
                TokenForUser = authResult.AccessToken;
            }
            catch (Exception)
            {
                var securePassword = new SecureString();
                foreach (char c in password)
                    securePassword.AppendChar(c);

                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    //          authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);
                    authResult = await IdentityClientApp.AcquireTokenByUsernamePasswordAsync(Scopes, username, securePassword);
                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }
            return TokenForUser;
        }

        static DateTimeOffset Expiration;
        static GraphServiceClient graphClient;
        static string TokenForUser;

       // static string clientIdForUser = Configuration.ClientIdForUserAuthn;
        
        
    
        static void Main(string[] args)
        {


            if (!System.IO.Directory.Exists("Output"))
            {
                System.IO.Directory.CreateDirectory("Output");
            }

            var Tracefile = "Output\\Trace.log";

            if (System.IO.File.Exists(Tracefile))
            {
                System.IO.File.Delete(Tracefile);
            }

            Trace.Listeners.Add(new TextWriterTraceListener(Tracefile));
            Trace.AutoFlush = true;
            Trace.Flush();





             
             

        Trace.WriteLine("Starting program.....");

            Read_Config_Files();
            client = GetAuthenticatedClientForUser(Configuration.username,Configuration.password);

            Trace.Write("Getting License Information from tenant......");
            Get_Licenses_SKU(client);
            Trace.WriteLine("DONE");

            Trace.Write("Getting Users License Status from tenant......");
            Get_Users_Status(client);
            Trace.WriteLine("DONE");

            Generate_End_User_Report();

            Evaluate_Changes();

            if (!Configuration.LogOnlyMode) { 
                foreach (Actions_Class action in Actions_to_do)
                {

                    if (action.Action == "ADD")
                    {
                        var response = Add_LicenseAsync(action, TokenForUser);
                        Trace.WriteLine("Added licence (" + action.Profile + ") for " + action.UserPrincipalName);
                    }

                    if (action.Action == "REMOVE")
                    {
                        var response = Remove_LicenseAsync(action, TokenForUser);
                        Trace.WriteLine("Removed all licences for " + action.UserPrincipalName);
                    }


                }
            } else
            {
                Trace.WriteLine("LoggingModeOnly is enabled, so I'll not do any change!");
            }



        }
    }
}
