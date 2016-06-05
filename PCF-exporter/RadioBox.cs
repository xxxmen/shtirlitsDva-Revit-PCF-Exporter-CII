using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PCF_Functions
{
    public partial class RadioBox : System.Windows.Forms.GroupBox
    {
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            var radioButton = e.Control as RadioButton;
            if (radioButton != null) radioButton.Click += radioButton_Click;
        }

        void radioButton_Click(object sender, EventArgs e)
        {
            var radio = (RadioButton)sender;
            if (!radio.Checked) radio.Checked = true;
        }

    }
}
