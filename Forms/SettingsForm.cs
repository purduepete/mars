using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MarsAddIn.HelperClasses;
using Newtonsoft.Json;
using json = MarsAddIn.HelperClasses.MarsSettings;
using helper = MarsAddIn.HelperClasses.HelperClass;

namespace MarsAddIn.Forms
{
    public class SettingsForm : Form
    {
        private ListBox _serverListBox;
        private TextBox _serverTextBox;
        private Button _addButton;
        private Button _removeButton;
        private Button _saveButton;
        private Button _cancelButton;
        private Label _lblServers;
        private Label _lblAdd;

        public SettingsForm()
        {
            Text = @"Add-In Settings";
            Width = 350;
            Height = 340;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            _lblServers = new Label
            {
                Text = @"Database Servers:",
                Top = 10,
                Left = 15,
                AutoSize = true
            };

            _serverListBox = new ListBox
            {
                Top = 30,
                Left = 15,
                Width = 305,
                Height = 180
            };

            _lblAdd = new Label
            {
                Text = @"Server Name:",
                Top = 218,
                Left = 15,
                AutoSize = true
            };

            _serverTextBox = new TextBox
            {
                Top = 236,
                Left = 15,
                Width = 185
            };

            _addButton = new Button
            {
                Text = @"Add",
                Top = 235,
                Left = 205,
                Width = 55
            };
            _addButton.Click += AddButton_Click;

            _removeButton = new Button
            {
                Text = @"Remove",
                Top = 235,
                Left = 265,
                Width = 55
            };
            _removeButton.Click += RemoveButton_Click;

            _saveButton = new Button
            {
                Text = @"Save",
                Top = 270,
                Left = 117,
                Width = 100
            };
            _saveButton.Click += SaveButton_Click;

            _cancelButton = new Button
            {
                Text = @"Cancel",
                Top = 270,
                Left = 220,
                Width = 100
            };
            _cancelButton.Click += (s, e) => Close();

            var _openFileButton = new Button
            {
                Text = @"Open Config",
                Top = 270,
                Left = 15,
                Width = 100
            };
            _openFileButton.Click += (s, e) =>
            {
                if (File.Exists(helper.SettingsPath))
                    System.Diagnostics.Process.Start(helper.SettingsPath);
            };

            Controls.AddRange(new Control[]
            {
                _lblServers, _serverListBox, _lblAdd, _serverTextBox,
                _addButton, _removeButton, _saveButton, _cancelButton, _openFileButton
            });

            LoadSettings();
        }

        private void LoadSettings()
        {
            if (!File.Exists(helper.SettingsPath)) return;

            var settings = JsonConvert.DeserializeObject<json>(File.ReadAllText(helper.SettingsPath));
            if (settings?.DatabaseServers == null) return;

            foreach (var server in settings.DatabaseServers)
                _serverListBox.Items.Add(server);
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            var server = _serverTextBox.Text.Trim();
            if (string.IsNullOrEmpty(server)) return;

            if (_serverListBox.Items.Contains(server))
            {
                MessageBox.Show(@"Server already exists.", @"Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _serverListBox.Items.Add(server);
            _serverTextBox.Clear();
            _serverTextBox.Focus();
        }

        private void RemoveButton_Click(object sender, EventArgs e)
        {
            if (_serverListBox.SelectedIndex < 0) return;
            _serverListBox.Items.RemoveAt(_serverListBox.SelectedIndex);
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            var dirLocation = Path.GetDirectoryName(helper.SettingsPath);
            if (!Directory.Exists(dirLocation))
                Directory.CreateDirectory(dirLocation);

            json settings;
            if (File.Exists(helper.SettingsPath))
                settings = JsonConvert.DeserializeObject<json>(File.ReadAllText(helper.SettingsPath)) ?? new json();
            else
                settings = new json();

            settings.DatabaseServers = _serverListBox.Items.Cast<string>().ToList();

            File.WriteAllText(helper.SettingsPath, JsonConvert.SerializeObject(settings, Formatting.Indented));

            MessageBox.Show(@"Settings saved.", @"Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }
    }
}
