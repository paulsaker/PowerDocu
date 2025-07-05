﻿using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using PowerDocu.AppDocumenter;
using PowerDocu.SolutionDocumenter;
using PowerDocu.Common;

namespace PowerDocu.GUI
{
    public partial class PowerDocuForm : Form
    {
        public PowerDocuForm()
        {
            InitializeComponent();
            LoadConfig();
            NotificationHelper.AddNotificationReceiver(
                new PowerDocuFormNotificationReceiver(appStatusTextBox)
            );
            using (var stream = File.OpenRead("Icons\\PowerDocu.ico"))
            {
                this.Icon = new Icon(stream);
            }
            InitialChecks();
        }

        private void LoadConfig()
        {
            if (configHelper == null)
                configHelper = new ConfigHelper();
            configHelper.LoadConfigurationFromFile();

            // Load existing settings
            outputFormatComboBox.SelectedItem = configHelper.outputFormat;
            documentChangesOnlyRadioButton.Checked = configHelper.documentChangesOnlyCanvasApps;
            documentEverythingRadioButton.Checked = !configHelper.documentChangesOnlyCanvasApps;
            documentDefaultsCheckBox.Checked = configHelper.documentDefaultValuesCanvasApps;
            documentSampleDataCheckBox.Checked = configHelper.documentSampleData;
            flowActionSortOrderComboBox.SelectedItem = configHelper.flowActionSortOrder;

            // Load newly added settings
            solutionCheckBox.Checked = configHelper.documentSolution;
            flowsCheckBox.Checked = configHelper.documentFlows;
            appsCheckBox.Checked = configHelper.documentApps;
            appPropertiesCheckBox.Checked = configHelper.documentAppProperties;
            variablesCheckBox.Checked = configHelper.documentAppVariables;
            dataSourcesCheckBox.Checked = configHelper.documentAppDataSources;
            resourcesCheckBox.Checked = configHelper.documentAppResources;
            controlsCheckBox.Checked = configHelper.documentAppControls;

            // Load Word template if available
            if (configHelper.wordTemplate != null)
            {
                openWordTemplateDialog.FileName = configHelper.wordTemplate;
                wordTemplateInfoLabel.Text =
                    "Template: " + Path.GetFileName(configHelper.wordTemplate);
            }
        }

        private async void InitialChecks()
        {
            //check for newer release
            if (await PowerDocuReleaseHelper.HasNewerPowerDocuRelease())
            {
                newReleaseButton.Visible = true;
                newReleaseLabel.Text += PowerDocuReleaseHelper.latestVersionTag;
                newReleaseLabel.Visible = true;
                string newReleaseMessage =
                    $"A new PowerDocu release has been found: {PowerDocuReleaseHelper.latestVersionTag}";
                NotificationHelper.SendNotification(newReleaseMessage);
                NotificationHelper.SendNotification(
                    "Please visit "
                        + PowerDocuReleaseHelper.latestVersionUrl
                        + " or press the Update button to download it"
                );
                NotificationHelper.SendNotification(Environment.NewLine);
                statusLabel.Text = newReleaseMessage;
            }
            //check for number of files
            int connectorIcons = ConnectorHelper.numberOfConnectorIcons();
            if (connectorIcons < 100)
            {
                NotificationHelper.SendNotification(
                    $"Only {connectorIcons} connector icons were found. Please update the Connectors list (press the Green Cloud Download icon)"
                );
            }
        }

        private void SelectZIPFileButton_Click(object sender, EventArgs e)
        {
            if (openFileToParseDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilesToDocumentLabel.Text =
                    "Select either the full documentation or just the image generation to start the documentation process for the following selected files:"
                    + Environment.NewLine;
                foreach (string fileName in openFileToParseDialog.FileNames)
                {
                    selectedFilesToDocumentLabel.Text +=
                        "   " + Path.GetFileName(fileName) + Environment.NewLine;
                }
                startDocumentationButton.Visible = true;
                startImageGenerationButton.Visible = true;
            }
        }

        private void SelectWordTemplateButton_Click(object sender, EventArgs e)
        {
            if (openWordTemplateDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    clearWordTemplateButton.Visible = true;
                    wordTemplateInfoLabel.Text =
                        "Template: " + Path.GetFileName(openWordTemplateDialog.FileName);
                    NotificationHelper.SendNotification(
                        "Selected Word template " + openWordTemplateDialog.FileName
                    );
                }
                catch (Exception ex)
                {
                    NotificationHelper.SendNotification("An error has occurred:");
                    NotificationHelper.SendNotification("Error message: " + ex.Message);
                    NotificationHelper.SendNotification(Environment.NewLine);
                    NotificationHelper.SendNotification("Details:");
                    NotificationHelper.SendNotification(ex.StackTrace);
                    NotificationHelper.SendNotification(Environment.NewLine);
                }
            }
        }

        private void NewReleaseButton_Click(object sender, EventArgs e)
        {
            var sInfo = new System.Diagnostics.ProcessStartInfo(
                PowerDocuReleaseHelper.latestVersionUrl
            )
            {
                UseShellExecute = true,
            };
            System.Diagnostics.Process.Start(sInfo);
        }

        private async void UpdateConnectorIconsButton_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "Updating connector icons...";
            statusLabel.Refresh();
            updateConnectorIconsButton.Enabled = false;
            updateConnectorIconsButton.IconColor = Color.DarkGray;
            await ConnectorHelper.UpdateConnectorIcons();
            updateConnectorIconsButton.Enabled = true;
            updateConnectorIconsButton.IconColor = Color.Green;
            updateConnectorIconsLabel.Text =
                $"Update your existing set of connector icons\n({ConnectorHelper.numberOfConnectors()} connectors, {ConnectorHelper.numberOfConnectorIcons()} icons)";
            statusLabel.Text =
                $"Connector icons have been updated ({ConnectorHelper.numberOfConnectors()} connectors, {ConnectorHelper.numberOfConnectorIcons()} icons)";
        }

        private void SyncConfigHelper()
        {
            if (configHelper == null)
            {
                configHelper = new ConfigHelper();
            }
            configHelper.outputFormat = outputFormatComboBox.SelectedItem.ToString();
            configHelper.documentChangesOnlyCanvasApps = documentChangesOnlyRadioButton.Checked;
            configHelper.documentDefaultValuesCanvasApps = documentDefaultsCheckBox.Checked;
            configHelper.documentSampleData = documentSampleDataCheckBox.Checked;
            configHelper.flowActionSortOrder = flowActionSortOrderComboBox.SelectedItem.ToString();
            configHelper.wordTemplate = openWordTemplateDialog.FileName;
            configHelper.documentSolution = solutionCheckBox.Checked;
            configHelper.documentFlows = flowsCheckBox.Checked;
            configHelper.documentApps = appsCheckBox.Checked;
            configHelper.documentAppProperties = appPropertiesCheckBox.Checked;
            configHelper.documentAppVariables = variablesCheckBox.Checked;
            configHelper.documentAppDataSources = dataSourcesCheckBox.Checked;
            configHelper.documentAppResources = resourcesCheckBox.Checked;
            configHelper.documentAppControls = controlsCheckBox.Checked;
        }

        private async void SaveConfigButton_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "Saving configuration...";
            statusLabel.Refresh();
            SyncConfigHelper();
            configHelper.SaveConfigurationToFile();
            statusLabel.Text = "New default configuration has been saved.";
        }

        private async void StartDocumentationButton_Click(object sender, EventArgs e)
        {
            startDocumentation(true);
        }

        private async void StartImageGenerationButton_Click(object sender, EventArgs e)
        {
            startDocumentation(false);
        }

        //fullDocumentation = true to start full documentation generation , false to start image generation
        private void startDocumentation(bool fullDocumentation = true)
        {
            SyncConfigHelper();
            statusLabel.Text =
                $"Starting documentation process for {openFileToParseDialog.FileNames.Length} files...";
            statusLabel.Refresh();
            startDocumentationButton.Enabled = false;
            startImageGenerationButton.Enabled = false;
            startDocumentationButton.IconColor = Color.DarkGray;
            startImageGenerationButton.IconColor = Color.DarkGray;
            foreach (string fileName in openFileToParseDialog.FileNames)
            {
                try
                {
                    NotificationHelper.SendNotification(
                        "Preparing to parse file " + fileName + ", please wait."
                    );
                    Cursor = Cursors.WaitCursor; // change cursor to hourglass type
                    if (fileName.EndsWith(".zip"))
                    {
                        NotificationHelper.SendNotification(
                            "Trying to process Solution, Apps, and Flows"
                        );
                        SolutionDocumentationGenerator.GenerateDocumentation(
                            fileName,
                            fullDocumentation,
                            configHelper
                        );
                    }
                    else if (fileName.EndsWith(".msapp"))
                    {
                        AppDocumentationGenerator.GenerateDocumentation(
                            fileName,
                            fullDocumentation,
                            configHelper
                        );
                    }
                    NotificationHelper.SendNotification("Documentation generation completed.");
                    statusLabel.Text = $"Documentation process completed";
                }
                catch (Exception ex)
                {
                    statusLabel.Text = $"An error occured. Please check the log for details.";
                    MessageBox.Show(
                        $"An error has occurred.\n\nError message: {ex.Message}\n\n"
                            + $"Details:\n\n{ex.StackTrace}"
                    );
                }
                finally
                {
                    NotificationHelper.SendNotification(Environment.NewLine);
                    Cursor = Cursors.Arrow; // change cursor to normal type
                    startDocumentationButton.Enabled = true;
                    startImageGenerationButton.Enabled = true;
                    startDocumentationButton.IconColor = Color.Green;
                    startImageGenerationButton.IconColor = Color.Green;
                }
            }
        }

        private async void ClearWordTemplateButton_Click(object sender, EventArgs e)
        {
            openWordTemplateDialog.FileName = "";
            clearWordTemplateButton.Visible = false;
            wordTemplateInfoLabel.Text = "No Word template selected";
        }

        private void SizeChangedHandler(object sender, EventArgs e)
        {
            appStatusTextBox.Size = new Size(
                ClientSize.Width - convertToDPISpecific(40),
                ClientSize.Height - convertToDPISpecific(100)
            );
            dynamicTabControl.Height = ClientSize.Height - convertToDPISpecific(50);
            dynamicTabControl.Width = ClientSize.Width;
            statusLabel.Width = convertToDPISpecific(ClientSize.Width - convertToDPISpecific(20));

            generateDocuPanel.Size = new Size(
                ClientSize.Width - convertToDPISpecific(30),
                ClientSize.Height - convertToDPISpecific(25)
            );
        }

        private void DpiChangedHandler(object sender, EventArgs e)
        {
            MinimumSize = new Size(convertToDPISpecific(800), convertToDPISpecific(350));
        }

        private void OutputFormatComboBox_Changed(object sender, EventArgs e)
        {
            if (outputFormatComboBox != null && selectWordTemplateButton != null)
            {
                if (
                    outputFormatComboBox.SelectedItem.ToString().Equals(OutputFormatHelper.Word)
                    || outputFormatComboBox.SelectedItem.ToString().Equals(OutputFormatHelper.All)
                )
                {
                    selectWordTemplateButton.Enabled = true;
                    wordTemplateInfoLabel.ForeColor = Color.Black;
                }
                else
                {
                    selectWordTemplateButton.Enabled = false;
                    wordTemplateInfoLabel.ForeColor = Color.Gray;
                }
            }
        }
    }

    public class PowerDocuFormNotificationReceiver : NotificationReceiverBase
    {
        private readonly TextBox notificationTextBox;

        public PowerDocuFormNotificationReceiver(TextBox textBox)
        {
            notificationTextBox = textBox;
        }

        public override void Notify(string notification)
        {
            notificationTextBox.AppendText(notification);
            notificationTextBox.AppendText(Environment.NewLine);
        }
    }
}
