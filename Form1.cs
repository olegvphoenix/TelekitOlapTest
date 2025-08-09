using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Telerik.Pivot.Adomd;
using Telerik.Pivot.Core;
using Telerik.Pivot.Core.Filtering;
using Telerik.Pivot.Core.Olap;
using Telerik.WinControls.Export;

namespace OlapTest
{
    /// <summary>
    /// https://corp.etsp.ru/company/personal/user/7120/tasks/task/view/396656/?ta_sec=tasks&ta_sub=list&ta_el=title_click
    /// </summary>
    public partial class Form1 : Form
    {
        /// <summary>
        /// Constant for the date dimension name in the cube
        /// </summary>
        private const string DATE_DIMENSION_NAME = "[Дата-Календарь].[Дата]";

        private Stopwatch stopwatch = Stopwatch.StartNew();
        private DataProviderStatus _lastStatus = DataProviderStatus.Uninitialized;
        private AdomdConnection _monitoredConnection;
        private static Form1 _activeInstance;

        /// <summary>
        /// Start date of the period
        /// </summary>
        public DateTime DateFrom => _datePickerFrom.Value;

        /// <summary>
        /// End date of the period
        /// </summary>
        public DateTime DateTo => _datePickerTo.Value;

        public Form1()
        {
            InitializeComponent();

            // Save reference to active form for global exception logging
            _activeInstance = this;

            // Initialize logging to text field
            LogMessage("OLAP Test application initialization");
            LogMessage("Global exception handling activated");

            _autoSaveLayout.Checked = Properties.Settings.Default.ExportXmlProfile;

            // Subscribe to date change events

            SetupAdomdProvider();

            // Add context menu for exception testing
            SetupTestExceptionMenu();

            // Initialize export XML checkbox
            _exportXmlCheckBox.Checked = Properties.Settings.Default.ExportXmlProfile;
        }



        /// <summary>
        /// Setup context menu for exception testing
        /// </summary>
        private void SetupTestExceptionMenu()
        {
            try
            {
                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add(new ToolStripMenuItem("Clear Log", null, (s, e) => ClearLog()));

                _logTextBox.ContextMenuStrip = contextMenu;

                LogMessage("Test context menu added (RMB on log panel)");
            }
            catch (Exception ex)
            {
                LogMessage($"Test menu setup error: {ex.Message}");
            }
        }

        /// <summary>
        /// Clear log panel
        /// </summary>
        private void ClearLog()
        {
            try
            {
                _logTextBox.Clear();
                LogMessage("=== LOG CLEARED ===");
                LogMessage("Log panel ready");
            }
            catch (Exception ex)
            {
                LogMessage($"Log clear error: {ex.Message}");
            }
        }

        /// <summary>
        /// Thread-safe logging of messages to text field
        /// </summary>
        /// <param name="message">Message to log</param>
        private void LogMessage(string message)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";

                if (_logTextBox.InvokeRequired)
                {
                    _logTextBox.BeginInvoke((MethodInvoker)delegate
                    {
                        AppendLogMessage(logEntry);
                    });
                }
                else
                {
                    AppendLogMessage(logEntry);
                }
            }
            catch (Exception ex)
            {
                // If we can't write to log, ignore the error
                System.Diagnostics.Debug.WriteLine($"Logging error: {ex.Message}");
            }
        }

        /// <summary>
        /// Add message to text field with auto-scroll
        /// </summary>
        private void AppendLogMessage(string logEntry)
        {
            _logTextBox.AppendText(logEntry + Environment.NewLine);
            _logTextBox.SelectionStart = _logTextBox.Text.Length;
            _logTextBox.ScrollToCaret();

            // Limit log lines count (last 1000 lines)
            var lines = _logTextBox.Lines;
            if (lines.Length > 1000)
            {
                var newLines = lines.Skip(lines.Length - 1000).ToArray();
                _logTextBox.Lines = newLines;
            }
        }

        /// <summary>
        /// Static method for logging global exceptions
        /// </summary>
        /// <param name="source">Exception source</param>
        /// <param name="exception">Exception</param>
        public static void LogGlobalException(string source, Exception exception)
        {
            try
            {
                if (_activeInstance != null)
                {
                    _activeInstance.LogException(source, exception);
                }
            }
            catch
            {
                // If we can't log to form, ignore
            }
        }

        /// <summary>
        /// Log exceptions with detailed formatting
        /// </summary>
        /// <param name="source">Exception source</param>
        /// <param name="exception">Exception</param>
        private void LogException(string source, Exception exception)
        {
            try
            {
                LogMessage("═══════════════════════════════════════");
                LogMessage($"[EXCEPTION] {source}");
                LogMessage("═══════════════════════════════════════");

                var currentException = exception;
                int level = 0;

                while (currentException != null)
                {
                    var indent = new string(' ', level * 2);

                    LogMessage($"{indent}[TYPE] {currentException.GetType().Name}");
                    LogMessage($"{indent}[MESSAGE] {currentException.Message}");

                    if (!string.IsNullOrEmpty(currentException.Source))
                    {
                        LogMessage($"{indent}[SOURCE] {currentException.Source}");
                    }

                    if (currentException.TargetSite != null)
                    {
                        LogMessage($"{indent}[METHOD] {currentException.TargetSite.DeclaringType?.Name}.{currentException.TargetSite.Name}");
                    }

                    if (!string.IsNullOrEmpty(currentException.StackTrace))
                    {
                        LogMessage($"{indent}[CALL STACK]");
                        var stackLines = currentException.StackTrace.Split('\n');
                        foreach (var line in stackLines.Take(5)) // Show first 5 stack lines
                        {
                            var trimmedLine = line.Trim();
                            if (!string.IsNullOrEmpty(trimmedLine))
                            {
                                LogMessage($"{indent}   {trimmedLine}");
                            }
                        }

                        if (stackLines.Length > 5)
                        {
                            LogMessage($"{indent}   ... and {stackLines.Length - 5} more lines");
                        }
                    }

                    // Additional information for specific exception types
                    LogSpecificExceptionDetails(currentException, indent);

                    currentException = currentException.InnerException;
                    level++;

                    if (currentException != null)
                    {
                        LogMessage($"{indent}[INNER EXCEPTION]");
                    }
                }

                LogMessage("═══════════════════════════════════════");
            }
            catch (Exception logEx)
            {
                // If error in logging itself, write minimal information
                LogMessage($"[LOGGING ERROR] {logEx.Message}");
                LogMessage($"[ORIGINAL EXCEPTION] {exception?.Message ?? "null"}");
            }
        }

        /// <summary>
        /// Log specific details for different exception types
        /// </summary>
        private void LogSpecificExceptionDetails(Exception exception, string indent)
        {
            try
            {
                switch (exception)
                {
                    case Microsoft.AnalysisServices.AdomdClient.AdomdException adomdEx:
                        LogMessage($"{indent}[ADOMD ERROR] {adomdEx.Message}");
                        if (!string.IsNullOrEmpty(adomdEx.Source))
                        {
                            LogMessage($"{indent}[ADOMD SOURCE] {adomdEx.Source}");
                        }
                        break;

                    case FileNotFoundException fileEx:
                        LogMessage($"{indent}[FILE NOT FOUND] {fileEx.FileName}");
                        break;


                    case ArgumentException argEx:
                        LogMessage($"{indent}[PARAMETER NAME] {argEx.ParamName ?? "Unknown"}");
                        break;

                    case InvalidOperationException _:
                        LogMessage($"{indent}[TYPE] Invalid operation");
                        break;

                    case NotImplementedException _:
                        LogMessage($"{indent}[TYPE] Function not implemented");
                        break;

                    case UnauthorizedAccessException _:
                        LogMessage($"{indent}[TYPE] Access denied");
                        break;

                    case TimeoutException timeoutEx:
                        LogMessage($"{indent}[TYPE] Timeout exceeded");
                        break;

                    case System.Security.SecurityException secEx:
                        LogMessage($"{indent}[SECURITY ERROR] {secEx.Message}");
                        if (!string.IsNullOrEmpty(secEx.PermissionType?.Name))
                        {
                            LogMessage($"{indent}[PERMISSION TYPE] {secEx.PermissionType.Name}");
                        }
                        break;

                    case OutOfMemoryException _:
                        LogMessage($"{indent}[CRITICAL ERROR] Out of memory");
                        break;

                    case StackOverflowException _:
                        LogMessage($"{indent}[CRITICAL ERROR] Stack overflow");
                        break;
                }
            }
            catch
            {
                // Ignore errors in detailed logging
            }
        }

        private string getSessionApplicationName()
        {
            return $"OLAP_TEST_{GetCurrentUser().Replace('\\', '_')}";
        }

        private void SetupAdomdProvider()
        {
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

                AdomdDataProvider provider = new AdomdDataProvider();
                AdomdConnectionSettings settings = new AdomdConnectionSettings();

                settings.ConnectionString = $"Data Source=srvt-sqlolap-01;Catalog=Remains_t;Cube=Остатки;Application Name={getSessionApplicationName()};";
                settings.Cube = "Остатки";
                settings.Database = "Remains_t";

                provider.ConnectionSettings = settings;

                this.radPivotGrid1.PivotGridElement.DataProvider = provider;
                LoadPivotConfiguration();

                provider.StatusChanged += (s, e) =>
                {
                    var elapsedMs = stopwatch.ElapsedMilliseconds;
                    _lastStatus = e.NewStatus;

                    if (e.NewStatus == DataProviderStatus.RetrievingData)
                    {
                        SavePivotConfiguration();
                        stopwatch.Restart();
                        StartLongOperationAsync();
                    }
                    else if (e.NewStatus == DataProviderStatus.Ready)
                    {
                        stopwatch.Stop();
                    }
                    else if (e.NewStatus == DataProviderStatus.Faulted)
                    {
                        LogMessage($"[UPDATE ERROR] {DataProviderStatus.Faulted}, ex={e.Error.Message}, det={e.Error.InnerException.Message}");
                        MessageBox.Show($"Error loading configuration: {e.Error.InnerException.Message}",
                        "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Error);

                        stopwatch.Stop();
                    }

                        UpdateLoadStatus(e.NewStatus.ToString(), elapsedMs, e.NewStatus == DataProviderStatus.Ready);
                };

                RefreshData();
                var providerState = provider.Status;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}",
                "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error);
            }
        }

        public void UpdateLoadStatus(string message, long elapsedMs, bool writeToLog)
        {
            if (_statusStrip.InvokeRequired)
            {
                _statusStrip.BeginInvoke((MethodInvoker)delegate
                {
                    UpdateLoadStatus(message, elapsedMs, writeToLog);
                });
                return;
            }

            _statusLabel.Text = $"{message}: {TimeSpan.FromMilliseconds(elapsedMs).TotalSeconds} sec";
            if (writeToLog)
                LogMessage($"[UPDATE] {_statusLabel.Text}");
        }

        private async void StartLongOperationAsync()
        {
            await Task.Run(() =>
            {
                while (_lastStatus == DataProviderStatus.RetrievingData)
                {
                    Thread.Sleep(50);
                    UpdateLoadStatus(_lastStatus.ToString(), stopwatch.ElapsedMilliseconds, false);
                }
            });
        }


        private void SavePivotConfiguration()
        {
            if (!Properties.Settings.Default.ExportXmlProfile)
                return;

            // Check if operation needs to be executed in UI thread
            if (radPivotGrid1.InvokeRequired)
            {
                radPivotGrid1.BeginInvoke((MethodInvoker)delegate
                {
                    SavePivotConfiguration();
                });
                return;
            }

            try
            {
                using (var stringWriter = new System.IO.StringWriter())
                {
                    // Create XmlWriter with settings
                    var settings = new System.Xml.XmlWriterSettings
                    {
                        Indent = true,          // Formatting with indents
                        Encoding = Encoding.UTF8 // UTF-8 encoding
                    };

                    using (var xmlWriter = System.Xml.XmlWriter.Create(stringWriter, settings))
                    {
                        // Save layout to XmlWriter (now safe in UI thread)
                        radPivotGrid1.SaveLayout(xmlWriter);

                        // Mandatory call to Flush to complete writing
                        xmlWriter.Flush();

                        // Get XML as string
                        string pivotGridLayoutXml = stringWriter.ToString();

                        // Now we can save the string where needed:
                        // - To application settings
                        Properties.Settings.Default.PivotGridConfiguration = pivotGridLayoutXml;
                        Properties.Settings.Default.Save();

                        LogMessage("[CONFIG] Layout saved successfully");

                        // If XML export is enabled, also save to file
                        if (Properties.Settings.Default.ExportXmlProfile)
                        {
                            try
                            {
                                var fileName = $"PivotProfile_{DateTime.Now:yyyyMMdd_HHmmss}.xml";
                                var filePath = Path.Combine(Application.StartupPath, "Profiles", fileName);
                                
                                // Ensure directory exists
                                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                                
                                File.WriteAllText(filePath, pivotGridLayoutXml);
                                LogMessage($"[EXPORT] XML profile exported to: {filePath}");
                            }
                            catch (Exception ex)
                            {
                                LogMessage($"[EXPORT ERROR] Failed to export XML profile: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to save configuration: {ex.Message}");
            }
        }

        private void LoadPivotConfiguration()
        {
            if (!Properties.Settings.Default.ExportXmlProfile)
                return;

            // Check if operation needs to be executed in UI thread
            if (radPivotGrid1.InvokeRequired)
            {
                radPivotGrid1.BeginInvoke((MethodInvoker)delegate
                {
                    LoadPivotConfiguration();
                });
                return;
            }

            try
            {
                if (!string.IsNullOrEmpty(Properties.Settings.Default.PivotGridConfiguration))
                {
                    using (XmlReader reader = XmlReader.Create(new StringReader(Properties.Settings.Default.PivotGridConfiguration)))
                    {
                        // Restore configuration from XML (now safe in UI thread)
                        radPivotGrid1.LoadLayout(reader);
                        LogMessage("[CONFIG] Layout loaded successfully");
                    }
                }
                else
                {
                    LogMessage("[CONFIG] No saved configuration to load");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to load configuration: {ex.Message}");
            }
        }

        /// <summary>
        /// Update button click handler - refreshes pivot grid data
        /// </summary>
        private void _updateButton_Click(object sender, EventArgs e)
        {
            try
            {
                LogMessage("[UPDATE] Refreshing pivot grid data...");

                // Refresh the pivot grid data
                RefreshData();

                LogMessage("[UPDATE] Pivot grid data refreshed successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to refresh pivot grid data: {ex.Message}");
                LogException("Update Button", ex);
            }
        }

        private void _autoSaveLayout_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ExportXmlProfile = _autoSaveLayout.Checked;
            Properties.Settings.Default.Save();
        }

        private void _exportXmlCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ExportXmlProfile = _exportXmlCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void _saveProfileButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"PivotProfile_{DateTime.Now:yyyyMMdd_HHmmss}.xml";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (var stringWriter = new StringWriter())
                        {
                            var settings = new XmlWriterSettings
                            {
                                Indent = true,
                                Encoding = Encoding.UTF8
                            };

                            using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
                            {
                                radPivotGrid1.SaveLayout(xmlWriter);
                                xmlWriter.Flush();
                            }

                            File.WriteAllText(saveFileDialog.FileName, stringWriter.ToString());
                            LogMessage($"[PROFILE] Layout saved to: {saveFileDialog.FileName}");
                            MessageBox.Show($"Profile saved successfully to:\n{saveFileDialog.FileName}", 
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to save profile: {ex.Message}");
                MessageBox.Show($"Error saving profile: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void _loadProfileButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var xmlContent = File.ReadAllText(openFileDialog.FileName);
                        using (var stringReader = new StringReader(xmlContent))
                        using (var xmlReader = XmlReader.Create(stringReader))
                        {
                            radPivotGrid1.LoadLayout(xmlReader);
                        }

                        LogMessage($"[PROFILE] Layout loaded from: {openFileDialog.FileName}");
                        MessageBox.Show($"Profile loaded successfully from:\n{openFileDialog.FileName}", 
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to load profile: {ex.Message}");
                MessageBox.Show($"Error loading profile: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void _exportButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"PivotExport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var exporter = new PivotGridSpreadExport(this.radPivotGrid1)
                        {
                            ExportFormat = SpreadExportFormat.Xlsx
                        };
                        exporter.RunExport(saveFileDialog.FileName, new SpreadExportRenderer());

                        LogMessage($"[EXPORT] XLSX exported to: {saveFileDialog.FileName}");

                        try
                        {
                            var psi = new ProcessStartInfo
                            {
                                FileName = saveFileDialog.FileName,
                                UseShellExecute = true
                            };
                            Process.Start(psi);
                            LogMessage($"[EXPORT] Opened: {saveFileDialog.FileName}");
                        }
                        catch (Exception openEx)
                        {
                            LogMessage($"[EXPORT WARNING] Could not open file: {openEx.Message}");
                        }

                        MessageBox.Show($"Data exported successfully to:\n{saveFileDialog.FileName}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Failed to export XLSX: {ex.Message}");
                MessageBox.Show($"Error exporting: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Get current user name in various formats
        /// </summary>
        private string GetCurrentUser()
        {
            try
            {
                // Try to get user from connection string first
                string connectionUser = ExtractUserFromConnectionString(_monitoredConnection?.ConnectionString);
                if (!string.IsNullOrEmpty(connectionUser))
                {
                    LogMessage($"[DEBUG] User from connection string: {connectionUser}");
                    return connectionUser;
                }

                // Try Windows Identity
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                if (identity?.Name != null)
                {
                    LogMessage($"[DEBUG] User from Windows Identity: {identity.Name}");
                    return identity.Name;
                }

                // Fallback to Environment
                string envUser = Environment.UserDomainName + "\\" + Environment.UserName;
                LogMessage($"[DEBUG] User from Environment: {envUser}");
                return envUser;
            }
            catch (Exception ex)
            {
                LogMessage($"[DEBUG] Error getting current user: {ex.Message}");
                return Environment.UserName;
            }
        }

        /// <summary>
        /// Extract user ID from connection string
        /// </summary>
        private string ExtractUserFromConnectionString(string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString))
                return null;

            try
            {
                var parts = connectionString.Split(';');
                foreach (var part in parts)
                {
                    var keyValue = part.Split('=');
                    if (keyValue.Length == 2)
                    {
                        var key = keyValue[0].Trim().ToLower();
                        var value = keyValue[1].Trim();

                        if (key == "user id" || key == "uid" || key == "user" || key == "username")
                        {
                            return value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[DEBUG] Error extracting user from connection string: {ex.Message}");
            }

            return null;
        }

        private void RefreshData()
        {
            try
            {
                var provider = radPivotGrid1.PivotGridElement.DataProvider as AdomdDataProvider;
                if (provider == null)
                    return;

                LogMessage($"Applying date filter: {DateFrom:dd.MM.yyyy} - {DateTo:dd.MM.yyyy}");

                var filterInfos = provider.FilterDescriptions.ToList();

                // Remove existing date filter if it exists
                var existingDateFilter = filterInfos.FirstOrDefault(f =>
                    f.MemberName == DATE_DIMENSION_NAME);

                if (existingDateFilter != null)
                {
                    provider.FilterDescriptions.Remove(existingDateFilter);
                }

                // Создаём новый фильтр по дате с использованием OlapSetCondition и Items
                var dateFilter = new AdomdFilterDescription()
                {
                    MemberName = DATE_DIMENSION_NAME
                };
                
                var condition = new OlapSetCondition()
                {
                    Comparison = SetComparison.Includes
                };
                for (var date = _datePickerFrom.Value.Date; date <= _datePickerTo.Value.Date; date = date.AddDays(1))
                {
                    // Используем правильный формат дат для куба с временем
                    var mdxItem = $"{DATE_DIMENSION_NAME}.&[{date:yyyy-MM-ddTHH:mm:ss}]";
                    condition.Items.Add(mdxItem);
                }
                dateFilter.Condition = condition;

                // Add filter to collection
                provider.FilterDescriptions.Add(dateFilter);

                // Refresh data
                provider.Refresh();
            }
            catch (Exception ex)
            {
                LogMessage($"Error applying date filter: {ex.Message}");
                MessageBox.Show($"Error applying date filter: {ex.Message}",
                "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error);
            }
        }
    }
}
