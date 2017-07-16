﻿using System;
using System.Windows.Forms;

namespace NTR_Exporter
{
    //This class is introduced to fix the default behaviour of radiobuttons in RadioBox. I couldn't get the radiobuttons
    //to switch on and off correctly inside one RadioBox. I found this solution somewhere on the net.
    public partial class RadioBox : System.Windows.Forms.GroupBox
    {
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            if (e.Control is RadioButton radioButton) radioButton.Click += radioButton_Click;
        }

        void radioButton_Click(object sender, EventArgs e)
        {
            var radio = (RadioButton)sender;
            if (!radio.Checked) radio.Checked = true;
        }

    }
}
