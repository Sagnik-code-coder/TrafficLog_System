import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    type IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SubmissionReportWebPartStrings';
import SummaryReport from './components/SubmissionReport';
import { ISubmissionReportProps } from './components/ISubmissionReportProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
//import * as $ from 'jquery';

export interface ISummaryReportWebPartProps {
    description: string;
}

export default class SubmissionReportWebPart extends BaseClientSideWebPart<ISubmissionReportProps> {

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';




    public render(): void {
        const element: React.ReactElement<ISubmissionReportProps> = React.createElement(
            SummaryReport,
            {
                description: this.properties.description,
                isDarkTheme: this._isDarkTheme,
                environmentMessage: this._environmentMessage,
                hasTeamsContext: !!this.context.sdks.microsoftTeams,
                userDisplayName: this.context.pageContext.user.displayName
            }
        );

        ReactDom.render(element, this.domElement);
        this.domElement.innerHTML =
            `
                <head>
                    <style type="text/css">
                        .ms-webpart-titleText {
                            display:none;
                        }
                        #tabs .ui-tabs-panel {
                            padding: 24px 0;
                        }
                    </style>
                    <script type="text/javascript">
                        document.onreadystatechange = function () {
                            if (document.readyState !== "complete") {
                                document.querySelector(".body-content").style.visibility = "hidden";
                                document.querySelector( "#loader").style.visibility = "visible";
                            } else {
                                document.querySelector("#loader").style.display = "none";
                                document.querySelector(".body-content").style.visibility = "visible";
                            }
                        };
                    </script> 
                </head>
                <body>
                <div id="loader" class="centerloader"></div>
                <div class="body-content">
                    <div class="container-fluid">
                       
                        <div class="alert alert-info" tkey="UserInaccessibility" id="UserInaccessibilityAlert" style="display:none">The current user is not allowed to visit this page.</div>
                            
                            <div id="summary">
                                <div id="dealersummary">
                                    <h1 class="heading" tkey="TrafficWritesFSummaryReport">Traffic, Writes  - Model Summary Report</h1>
            
            
                                    <div style="width:100%">
                                        <table style="width:100%" cellpadding="10" cellspacing="10">
                                            <tr style="width:100%">
                                                <td style="width:50%">
                                                    <label for="RegionArea" class="col-sm-3 col-form-label dc_area" tkey="RegionArea">Region/Area/Volume Group:</label>
                                                    
                                                    <select class="form-control" id="Regions">
                                                        
                                                    </select>
                                                </td>
                                                <td>
                                                    <label for="DistrictArea" class="col-sm-3 col-form-label dc_area" tkey="DistrictArea">District:</label> 
                                                    
                                                    <select class="form-control valid" id="Districts"></select>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                    <div style="width:100%;padding-top:20px">
                                        <table style="width:100%" cellpadding="10" cellspacing="10">
                                            <tr style="width:100%">
                                                <td style="width:50%">
                                                    <label for="Dealers" class="col-sm-3 col-form-label dc_area" tkey="Dealers">Clubhouses:</label> 
                                                    
                                                    <select class="form-control valid" id="Dealers"></select>
                                                </td>
                                                <td>
                                                    <label for="DealerSubmissionLabel" class="col-sm-3 col-form-label dc_area" tkey="DealerSubmissionLabel">Clubhouse Submission:</label> 
                                                    
                                                    <select class="form-control" id="SummaryDealerSubmission">
                                                        <option value="All"  selected>All</option>
                                                        <option value="Submitted"   >Submitted</option>
                                                        <option value="NotSubmitted"  >Not Submitted</option>
                                                    </select>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                    <div style="width:100%">
                                        <table style="width:100%" cellpadding="10" cellspacing="10">
                                            <tr style="width:100%">
                                                <td style="width:60%">
                                                  
                                                    <label for="CurrentWeek" class="col-sm-3 col-form-label dc_area" tkey="CurrentWeekLabel" style="margin-top: 10px;">Current Week:</label>
            
                                                    <label id="currentWeek" class="col-form-label dc_area" style="margin-left: 137px;margin-top: 11px;"></label>
                                                </td>
                                                <td></td>
                                            </tr>
                                        </table>
                                    </div>
                                    &nbsp; &nbsp;&nbsp;
                                    <div id="TrafficSubmissionReport" style="width:100%">
                                        <table style="width:100%" cellpadding="10" cellspacing="10">
                                            <tr style="width:100%">
                                                <td style="width:50%">
                                                    <div class="btn-group pull-left">
                                                        <button type="button" class="btn btn-dark primary_btn" tkey="ExportToCSV" id="export">Export</button>
                                                    </div>
                                                </td>
                                                <td><p style="text-align:right">
                                                     <button type="button" id="btnSummaryFilter" class="btn btn-dark primary_btn" tkey="Filter" style="margin-top: 10px;">Filter</button>
               
                                                </p></td>
                                            </tr>
                                        </table>
                                    </div>
            
                                    <div id="SummaryReport" style="width:100%">
                                        <table id="SummaryTable" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0">
                                            <thead>
            
                                                <tr>
                                                    <th tkey="DealerCode">Clubhouse Code </th>
                                                    <th tkey="Status">Status</th>
                                                    <th tkey="Region">Region</th>
                                                    <th tkey="District">District</th>
                                                    <th tkey="Timestamp">Timestamp</th>
                                                </tr>
                                            </thead>
                                            <tbody></tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        
                    </div>
                </div>
                        
                </body>
        
        
            `
    }

    protected onInit(): Promise<void> {

        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/jsgrid.min.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/jsgrid-theme.min.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/bootstrap.min.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/ui.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/responsive.bootstrap.min.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/bootstrap.css');
        SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/Site.css');

        SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery-1.10.2.js', { globalExportsName: 'jquery' }).then(() => {
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/bootstrap.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/respond.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/Scripts.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/notify.min.js');

            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery.validate.min.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery.validate.unobtrusive.min.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jsgrid.min.js');
            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery-ui.js');

            SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/MEPReport.js');
            //SPComponentLoader.loadScript('https://2721r7.sharepoint.com/SiteAssets/SummaryReport/js/csvExport.js');
        });


        // SPComponentLoader.loadScript('https://2721r7.sharepoint.com/SiteAssets/SummaryReport/js/lang.js');


        //this.loadJSONData();
        // this.populateDealersSelect();



        return this._getEnvironmentMessage().then(message => {
            this._environmentMessage = message;
        });
    }



    private _getEnvironmentMessage(): Promise<string> {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(context => {
                    let environmentMessage: string = '';
                    switch (context.app.host.name) {
                        case 'Office': // running in Office
                            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                            break;
                        case 'Outlook': // running in Outlook
                            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                            break;
                        case 'Teams': // running in Teams
                        case 'TeamsModern':
                            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                            break;
                        default:
                            environmentMessage = strings.UnknownEnvironment;
                    }

                    return environmentMessage;
                });
        }

        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        this._isDarkTheme = !!currentTheme.isInverted;
        const {
            semanticColors
        } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }

    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
