using System;
using System.Windows.Forms;

namespace FocaExcelExport.Ui
{
    public class MainControl : UserControl
    {
        private readonly TabControl _tabs;
        private readonly TabPage _exportTab;
        private readonly TabPage _compareTab;
        private bool _initialized;

        public MainControl()
        {
            Dock = DockStyle.Fill;

            _tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Alignment = TabAlignment.Top,
                Multiline = false,
            };

            _exportTab = new TabPage("Exportar");
            _compareTab = new TabPage("Comprobar");

            _tabs.TabPages.Add(_exportTab);
            _tabs.TabPages.Add(_compareTab);

            Controls.Add(_tabs);

            Load += MainControl_Load;
        }

        private void MainControl_Load(object sender, EventArgs e)
        {
            if (_initialized) return;
            _initialized = true;

            // Exportar tab
            TryEmbedForm(() => new ExportDialog { Embedded = true }, _exportTab);

            // Comprobar tab
            TryEmbedForm(() => new CompareDialog { Embedded = true }, _compareTab);
        }

        private static void TryEmbedForm(Func<Form> factory, TabPage host)
        {
            try
            {
                host.Controls.Clear();

                var form = factory();
                form.TopLevel = false;
                form.Dock = DockStyle.Fill;
                form.FormBorderStyle = FormBorderStyle.None;
                form.ShowInTaskbar = false;
                host.Controls.Add(form);
                form.Show();
            }
            catch
            {
                // Evitar que la UI completa de FOCA falle por excepciones del plugin.
            }
        }
    }
}

