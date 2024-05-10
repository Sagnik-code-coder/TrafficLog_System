import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'TabLinkWebPartStrings';
import TabLink from './components/TabLink';
import { ITabLinkProps } from './components/ITabLinkProps';
import * as jQuery from 'jquery';
import 'jqueryui';


export interface ITabLinkWebPartProps {
  description: string;
}

export default class TabLinkWebPart extends BaseClientSideWebPart<ITabLinkWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITabLinkProps> = React.createElement(
      TabLink,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
    const tabsHTML: string =
      `
    <style type="text/css">
        .ms-webpart-titleText {
            display: none;
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
        <div id="tabs" style="">
        <div>
            <ul>
                <li><a href="#Detailtabs" onclick="clickTab('0')" tkey="TrafficDetail">Traffic Detail</a></li>
                <li><a href="#TrafficSummary" onclick="clickTab('1')" tkey="TrafficSummary">Traffic Summary</a></li>
            </ul>
            <div id="TrafficDetail" style="padding:0px !important">
            </div>
            <div id="TrafficSummary" style="padding:0px !important">
            </div>
        </div>
    </div>
            </div>
        </div>
    </div>
    <div class="body-content">
        <div class="container-fluid">
            <div id="Detailtabs" style="">
                
                <div class="alert alert-info" tkey="UserInaccessibility" id="UserInaccessibilityAlert" style="display:none">The current user is not allowed to visit this page.</div>
                
                <div id="detail">
                    <div id="dealersummary">
                        <div class="row" style="width: 100%;max-width:2000px; ">
                            <div class="span4">
                                <fieldset>
                                    <div class="control-group">

                                        <label for="RegionArea"  style="padding-left:15px" class="form-label dc_area" tkey="RegionArea">Region/Area/Volume Group:</label>
                                         <select style="width:140px !important" class="form-control  dc_area" id="Regions"></select>
                                    </div>

                                </fieldset>
                            </div>
                            <div class="span3">
                                <fieldset>
                                    <div class="control-group">
                                        <label for="DistrictArea" class="form-label dc_area" tkey="DistrictArea">District:</label>
                                        <select style="width:140px !important" class="form-control dc_area" id="Districts"></select>
                                    </div>
                                </fieldset>
                            </div>
                            <div class="span3">
                                <fieldset>
                                    <div class="control-group">
                                        <label for="Dealers" class="form-label dc_area" tkey="Dealers">Clubhouses:</label>
                                        <select style="width:140px !important" class="form-control dc_area" id="Dealers"></select>
                                    </div>
                                </fieldset>
                            </div>
                        </div>
                                      
                        <div class="row" style="width: 100%; padding-top:20px;">
 
                                <div class="col">
                                    <fieldset>
                                        <div class="control-group"  id="TrafficDetailReport" style="text-align: right;">
                                            <button type="button" id="btnMTD" class="btn btn-dark primary_btn">MTD</button>
                                            <button type="button" id="btnMTDTrend" class="btn btn-dark primary_btn">MTD Trend</button>
                                            <button type="button" id="btnWOW" class="btn btn-dark primary_btn">WOW</button>
                                            <button   type="button" id="btnMOM" class="btn btn-dark primary_btn">MOM</button>
                                        </div>
                                        
                                    </fieldset>
                                </div>
                            </div>

                        <div class="row" style="width: 100%; padding-top:20px;">
                            <div class="col">
                                <fieldset>

                                    <div class="control-group">
                                        <label for="CurrentWeek" class="form-label dc_area" tkey="CurrentWeekLabel">Current Week:</label>
                                        
                                        <label id="currentWeek" class="form-label dc_area"> </label>

                                    </div>

                                </fieldset>
                            </div>
                            <div class="col">
                                <fieldset>
                                    
                                    <div class="control-group" style="text-align: right;">
                                        <button type="button" class="btn btn-dark primary_btn" tkey="ExportToCSV" id="exportdealer">Export</button>
                                    </div>
                                </fieldset>
                            </div>
                        </div>
   
                        


                        
                       
                            <div id="MTDReport" style="width:100%">
                                <table id="MTDSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="">
                                    <thead>
                                        <tr><th rowspan="3" tkey="Model">Model</th><th colspan="15">MTD</th></tr>
                                        <tr><th colspan="7" tkey="CurrentWeek">Current Week</th><th colspan="9" tkey="CurrentMonthMTD">Current Month (MTD)</th></tr>
                                        <tr>
                                            <th key="Traffic" tkey="Traffic">Traffic </th>
                                            <th key="Writes" tkey="Writes"> Writes</th>
                                            <th key="Closing" tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            <th key="AvgTraffic" tkey="AvgTraffic"> Avg. Traffic</th>
                                            <th key="AvgWrites" tkey="AvgWrites"> Avg. Writes</th>
                                            <th key="AvgClosing" tkey="AvgClosing"> Avg. Closing%</th>
                                            <th key="Traffic" tkey="Traffic"> Traffic </th>
                                            <th key="Writes" tkey="Writes"> Writes</th>
                                            <th key="Closing" tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesTarget">Monthly Sales Target</th>
                                            <th tkey="Achievenment">% Achievement</th>
                                            <th key="AvgTraffic" tkey="AvgTraffic">Avg. Traffic</th>
                                            <th key="AvgWrites" tkey="AvgWrites">Avg. Writes</th>
                                            <th key="AvgClosing" tkey="AvgClosing">Avg. Closing%</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                            <div id="MTDTrendReport" style="width:100%">
                                <table id="MTDTrend" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                    <thead>
                                        <tr><th rowspan="7" tkey="Model">Model</th><th colspan="20" tkey="MTDTrend">MTD Trend</th></tr>
                                        <tr><th colspan="4" tkey="Week1">Week 1</th><th colspan="4" tkey="Week2">Week 2</th><th colspan="4" tkey="Week3">Week 3</th><th colspan="4" tkey="Week4">Week 4</th><th colspan="4" tkey="Week5">Week 5</th></tr>
                                        <tr>
                                            <th tkey="Traffic"> Traffic </th>
                                            <th tkey="Writes"> Writes</th>
                                            <th tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            <th tkey="Traffic"> Traffic </th>
                                            <th tkey="Writes"> Writes</th>
                                            <th tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            <th tkey="Traffic"> Traffic </th>
                                            <th tkey="Writes"> Writes</th>
                                            <th tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            <th tkey="Traffic"> Traffic </th>
                                            <th tkey="Writes"> Writes</th>
                                            <th tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                            <th tkey="Traffic"> Traffic </th>
                                            <th tkey="Writes"> Writes</th>
                                            <th tkey="Closing"> Closing%</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                            <div id="WOWReport" style="width:100%">
                                <table id="WOWSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                    <thead>
                                        <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="WoWCOMPARISON">WoW COMPARISON</th><th colspan="6" tkey="WOWYoYCOMPARISON">Current Week YoY COMPARISON</th></tr>
                                        <tr>
                                            <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                            <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                            <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                            <th tkey="AreaAvgWrites">Area Avg. Writes %</th>
                                            <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                            <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                            <th tkey="MonthlySalesForecast">Monthly Sales Forecast %</th>
                                            <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                            <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                            <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                            <th tkey="AreaAvgWritesPercent">Area Avg. Writes %</th>
                                            <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                            <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                            <div id="MOMReport" style="width:100%">
                                <table id="MOMSubmission" class="table table-striped table-bordered dt-responsive nowrap data-entry" cellspacing="0" style="display:none;">
                                    <thead>
                                        <!--<tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISON">MTD YoY COMPARISON</th></tr>-->
                                        <tr><th rowspan="2" tkey="Model">Model</th><th colspan="7" tkey="MoMCOMPARISON">MOM COMPARISON</th><th colspan="6" tkey="YoYCOMPARISONDetailMOM">Current MTD YoY (Year-over-Year) COMPARISON</th></tr>
                                        <tr>
                                            <th tkey="DealerTrafficPercent">Clubhouse Traffic % </th>
                                            <th tkey="AreaAvgTrafficPercent">Area Avg. Traffic %</th>
                                            <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                            <th tkey="AreaAvgWrites">Area Avg. Writes %</th>
                                            <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                            <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                            <th tkey="MonthlySalesForecastPercent">Monthly Sales Forecast %</th>
                                            <th tkey="DealerTrafficPercent">Clubhouse Traffic %</th>
                                            <th tkey="AreaAvgTraffic">Area Avg. Traffic %</th>
                                            <th tkey="DealerWritesPercent">Clubhouse Writes %</th>
                                            <th tkey="AreaAvgWritesPercent">Area Avg. Writes %</th>
                                            <th tkey="DealerClosing">Clubhouse Closing (ppt)</th>
                                            <th tkey="AreaAvgClosing">Area Avg. Closing (ppt)</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>

                        </div>
                </div>

            </div>
        </div>
    </div>

</body>
<style type="text/css">
    .ms-webpart-titleText.ms-webpart-titleText, .ms-webpart-titleText > a {
        background-color: white;
        /*font-family: helvetica, verdana, arial, sans-serif, Geneva, sans-serif;*/
        font-size: 22px;
        font-weight: bold;
        color: #101010 !important;
        padding: 5px 15px;
        /*border: 2px solid #008AD2;*/
        box-shadow: none;
        margin-bottom: 8px;
        /*max-width: min-content;*/
        text-align: center !important;
        line-height:40px;
        /*border-bottom: 1px solid #e3e9ee;*/ 
    }

    .ms-WPBorderBorderOnly{

       border: 1px solid #fff;
    }
</style>
`
    this.domElement.innerHTML = tabsHTML;

    // Initialize jQuery tabs
    (jQuery("#tabs", this.domElement) as any).tabs();
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jsgrid/1.5.3/jsgrid.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jsgrid/1.5.3/jsgrid-theme.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.0.0-alpha.6/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css');
    SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.1.1/css/responsive.bootstrap.min.css');
    SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/bootstrap.css');
    SPComponentLoader.loadCss('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/css/Site.css');
    
    SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery-1.10.2.js')
        .then(() => {
            return SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jsgrid.min.js');
        })
        .then(() => {
            return SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/jquery-ui.js');

        })
        .then(() => {
          return SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/TrafficLog/js/DetailReport5.js');
        })

    
    
    
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
