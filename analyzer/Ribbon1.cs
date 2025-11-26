using Microsoft.Office.Tools.Ribbon;

namespace analyzer
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("ボタンが押下されました");
        }
    }
}
