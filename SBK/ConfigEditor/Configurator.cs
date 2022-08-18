using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.CSharp;

namespace ConfigEditor
{
    public partial class Configurator : Form
    {

        public Configurator()
        {
            InitializeComponent();
        }

        private void btOpen_Click(object sender, EventArgs e)
        {
            Assembly ass = null;
            listView.Items.Clear();
            if (openFileDialog.ShowDialog(this) != System.Windows.Forms.DialogResult.OK) return;
            try
            {
                string tmpFile = Path.GetTempFileName();
                File.Copy(openFileDialog.FileName, tmpFile, true);
                ass = Assembly.LoadFile(tmpFile);
                textBox2.Text = ass.GetName().Version.ToString();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(openFileDialog.FileName);
                textBox1.Text = fvi.ProductVersion;
            }
            catch (Exception)
            {
                MessageBox.Show("Возникла ошибка при открытии библиотеки.");
                textBox1.Text =
                    textBox2.Text = "";
                return;
            }

            Type t = ass.GetType("Config.Values", false, true);
            if (t == null)
            {
                MessageBox.Show("Не найден класс Config.Values.");
                textBox1.Text =
                    textBox2.Text = "";
                return;
            }


            foreach (var p in t.GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance))
            {

                if (p.FieldType.IsPrimitive || p.FieldType == typeof(String))
                    listView.Items.Add(new ListViewItem(new string[] { p.GetValue(null).ToString(), p.FieldType.FullName, p.Name }));
            }

        }

        private void btSaveAs_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("using System;");
                sb.AppendLine("using System.Reflection;");
                sb.AppendLine("[assembly: AssemblyVersion(\"" + textBox2.Text + "\")]");
                sb.AppendLine("[assembly: AssemblyInformationalVersion(\"" + textBox1.Text + "\")]");
                sb.AppendLine("namespace Config { public static class Values {");
                foreach (ListViewItem li in listView.Items)
                {
                    var t = Type.GetType(li.SubItems[1].Text);

                    if (t == typeof(String))
                        sb.AppendFormat("public static {0} {2} = @\"{1}\";\n", li.SubItems[1].Text, li.Text.Replace("\"", "\"\""), li.SubItems[2].Text);
                    else if (t == typeof(Boolean))
                    {
                        var val = Convert.ChangeType(li.Text, t);
                        sb.AppendFormat("public static {0} {2} = {1};\n", li.SubItems[1].Text, Convert.ToString(val).ToLower(), li.SubItems[2].Text);
                    }
                    else
                    {
                        var val = Convert.ChangeType(li.Text, t);
                        sb.AppendFormat("public static {0} {2} = {1};\n", li.SubItems[1].Text, Convert.ToString(val), li.SubItems[2].Text);
                    }
                }
                sb.Append(@"
                    }
                }");
                bool a = true;
                CSharpCodeProvider provider = new CSharpCodeProvider();
                CompilerParameters parameters = new CompilerParameters();
                // True - memory generation, false - external file generation
                parameters.OutputAssembly = saveFileDialog.FileName;
                // True - exe file generation, false - dll file generation
                parameters.GenerateExecutable = false;


                CompilerResults results = provider.CompileAssemblyFromSource(parameters, sb.ToString());

                if (results.Errors.Count > 0) throw new ApplicationException();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Возникла ошибка при сохранении файла.");
            }

        }

        private void listView_DoubleClick(object sender, EventArgs e)
        {
            listView.SelectedItems[0].BeginEdit();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }


        private void listView_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {

            string t = listView.Items[e.Item].SubItems[1].Text;
            try
            {
                Convert.ChangeType(e.Label, Type.GetType(t));
            }
            catch
            {
                e.CancelEdit = true;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Configurator_Load(object sender, EventArgs e)
        {

        }
    }
}
