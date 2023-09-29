<template>
  <div>
    <PowerBIVisualEmbedComponent
      :embed-config="sampleReportConfig"
      :phased-embedding="phasedEmbeddingFlag"
      :css-class-name="reportClass"
      :event-handlers="eventHandlersMap"
      @report-obj="setReportObj"
      v-if="isEmbedded"
    />
  </div>
</template>

<script>
import { models } from 'powerbi-client';

import 'powerbi-report-authoring';

import { PowerBIVisualEmbedComponent } from 'powerbi-client-vue-js';

// Flag which specifies whether to use phase embedding or not
const phasedEmbeddingFlag = false;

// CSS Class to be passed to the wrapper
const reportClass = 'report-container';

let report = null

export default {
  name: 'PowerBiReport',

  props: {

  },

  components: {
    PowerBIVisualEmbedComponent
  },

  data() {
    return {
      reportUrl: 'https://aka.ms/CaptureViewsReportEmbedConfig',
      //reportUrl: 'https://app.powerbi.com/groups/me/reports/a69a87af-7f5c-452b-8d86-50a01d185d46/ReportSection?ctid=74fcfbe3-83b9-40e9-855d-22269b17d96b&experience=power-bi',
      isEmbedded: false,

      // Overall status message of embedding
      displayMessage: 'The report is bootstrapped. Click Embed Report button to set the access token.',


      // Pass the basic embed configurations to the wrapper to bootstrap the report on first load
      // Values for properties like embedUrl, accessToken and settings will be set on click of button
      sampleReportConfig: {
        type: 'tile',
        embedUrl: undefined,
        tokenType: models.TokenType.Embed,
        accessToken: undefined,
        settings: undefined,
      },

      
      eventHandlersMap: new Map([
        ['loaded', () => console.log('Report has loaded')],
        ['rendered', () => console.log('Report has rendered')],
        ['error', (event) => {
            if (event) {
              console.error(event.detail);
            }
          },
        ],
        ['visualClicked', () => console.log('visual clicked')],
        ['pageChanged', (event) => console.log(event)],
      ]),

      // Store Embed object from Report component
      report,
      reportClass,
      phasedEmbeddingFlag
    }
  },

  created() {
    this.embedReport()
  },

  methods: {
     async embedReport() {
      console.log('Embed Report clicked')

      // Get the embed config from the service and set the reportConfigResponse
      const reportConfigResponse= await fetch(this.reportUrl);
      if (!reportConfigResponse?.ok) {
        console.error(`Failed to fetch config for report. Status: ${reportConfigResponse.status} ${reportConfigResponse.statusText}`);
        return;
      }

      const reportConfig = await reportConfigResponse.json();

      // Update the reportConfig to embed the PowerBI report
      this.sampleReportConfig = {
        ...this.sampleReportConfig,
        id: reportConfig.Id,
        embedUrl: reportConfig.EmbedUrl,
        accessToken: reportConfig.EmbedToken.Token
      };

      this.isEmbedded = true;

      // Update the display message
      this.displayMessage = 'Use the buttons above to interact with the report using Power BI Client APIs.';
    },

    /**
     * Change visual type
     *
     * @returns Promise<void>
     */
    async changeVisualType() {
      // Check Report is available or not
      if(!this.reportAvailable()) {
        return;
      }

      // Get all the pages of the report
      const pages = await this.report.getPages();

      // Check if the pages are available
      if (pages.length === 0) {
        this.displayMessage = 'No pages found.';
        return;
      }

      // Get active page of the report
      const activePage = pages.find((page) => page.isActive);

      if (!activePage) {
        this.displayMessage = 'No Active page found';
        return;
      }


      try {
        // Change the visual type using powerbi-report-authoring
        // For more information: https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/report-authoring-overview
        // Get the visual
        const visual = await activePage.getVisualByName('VisualContainer6');

        const response = await visual.changeType('lineChart');

        this.displayMessage = `The ${visual.type} was updated to lineChart.`;
        console.log(this.displayMessage);
        return response;
      } catch (error) {
        if (error === 'PowerBIEntityNotFound') {
          console.log('No Visual found with that name');
        } else {
          console.log(error);
        }
      }
    },

    /**
     * Hide Filter Pane
     *
     * @returns Promise<IHttpPostMessageResponse<void> | undefined>
     */
    async hideFilterPane() {
      // Check whether Report is available or not
      if(!this.reportAvailable()) {
        return;
      }

      // New settings to hide filter pane
      const settings = {
        panes: {
          filters: {
            expanded: false,
            visible: false,
          },
        },
      };

      try {
        const response = await this.report.updateSettings(settings);
        this.displayMessage = 'Filter pane is hidden.';
        console.log(this.displayMessage);
        return response;
      } catch (error) {
        console.error(error);
        return;
      }
    },

    /**
     * Set data selected event
     *
     * @returns void
     */
    setDataSelectedEvent() {
      this.eventHandlersMap = new Map(event),  () => ([
        ...this.eventHandlersMap,
        ['dataSelected', (event) => console.log(event)],
      ]);

      this.displayMessage = 'Data Selected event set successfully. Select data to see event in console.';
    },

    /**
     * Assign Embed Object from Report component to report
     * @param value
     */
    setReportObj(value) {
      this.report = value;
    },

    /**
     * Verify whether report is available or not
     */
    reportAvailable() {
      if (!this.report) {
        // Prepare status message for Error
        this.displayMessage = 'Report not available.';
        console.log(this.displayMessage);
        return false;
      }
      return true;
    }
  },

  

  
}
</script>


<style>
  .report-container {
    height: 75vh;
    margin: 8px auto;
    width: 90vw;
  }
</style>
