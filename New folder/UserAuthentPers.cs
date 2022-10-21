using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PersAhwal
{
    public partial class UserAuthentPers : UserControl
    {
        int AuthPersonlCount = 1;
        public UserAuthentPers()
        {
            InitializeComponent();
            PanelAuthPers.Height = 41;
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 1)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 2)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox16_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 3)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox18_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 4)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox20_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 5)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox23_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 6)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox25_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 7)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox27_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 8)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox29_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 9)
            {
                PanelAuthPers.Height += 41;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            if (AuthPersonlCount > 1)
            {
                txtAuthPerson1.Text = txtAuthPerson2.Text;
                txtAuthPerson2.Text = txtAuthPerson3.Text;
                txtAuthPerson3.Text = txtAuthPerson4.Text;
                txtAuthPerson4.Text = txtAuthPerson5.Text;
                txtAuthPerson5.Text = txtAuthPerson6.Text;
                txtAuthPerson6.Text = txtAuthPerson7.Text;
                txtAuthPerson7.Text = txtAuthPerson8.Text;
                txtAuthPerson8.Text = txtAuthPerson9.Text;
                txtAuthPerson9.Text = txtAuthPerson10.Text;
                txtAuthPerson10.Text = "";

                txtAuthPersonsex1.CheckState = txtAuthPersonsex2.CheckState;
                txtAuthPersonsex2.CheckState = txtAuthPersonsex3.CheckState;
                txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
                txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
                txtAuthPersonsex5.CheckState = txtAuthPersonsex6.CheckState;
                txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
                txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
                txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
                txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
                txtAuthPersonsex10.CheckState = CheckState.Unchecked;

                PanelAuthPers.Height -= 41;
                AuthPersonlCount -= 1;
            }
            else
            {
                txtAuthPerson1.Text = "";

                txtAuthPersonsex1.CheckState = CheckState.Unchecked;
                AuthPersonlCount = 1;
            }
        }

        private void pictureBox15_Click(object sender, EventArgs e)
        {

            txtAuthPerson2.Text = txtAuthPerson3.Text;
            txtAuthPerson3.Text = txtAuthPerson4.Text;
            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = txtAuthPerson6.Text;
            txtAuthPerson6.Text = txtAuthPerson7.Text;
            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex2.CheckState = txtAuthPersonsex3.CheckState;
            txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = txtAuthPersonsex6.CheckState;
            txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox17_Click(object sender, EventArgs e)
        {
            txtAuthPerson3.Text = txtAuthPerson4.Text;
            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = txtAuthPerson6.Text;
            txtAuthPerson6.Text = txtAuthPerson7.Text;
            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";

            txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = txtAuthPersonsex6.CheckState;
            txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox19_Click(object sender, EventArgs e)
        {

            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = txtAuthPerson6.Text;
            txtAuthPerson6.Text = txtAuthPerson7.Text;
            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = txtAuthPersonsex6.CheckState;
            txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox21_Click(object sender, EventArgs e)
        {

            txtAuthPerson5.Text = txtAuthPerson6.Text;
            txtAuthPerson6.Text = txtAuthPerson7.Text;
            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex5.CheckState = txtAuthPersonsex6.CheckState;
            txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox24_Click(object sender, EventArgs e)
        {

            txtAuthPerson6.Text = txtAuthPerson7.Text;
            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex6.CheckState = txtAuthPersonsex7.CheckState;
            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox26_Click(object sender, EventArgs e)
        {

            txtAuthPerson7.Text = txtAuthPerson8.Text;
            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex7.CheckState = txtAuthPersonsex8.CheckState;
            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox28_Click(object sender, EventArgs e)
        {

            txtAuthPerson8.Text = txtAuthPerson9.Text;
            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex8.CheckState = txtAuthPersonsex9.CheckState;
            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox30_Click(object sender, EventArgs e)
        {

            txtAuthPerson9.Text = txtAuthPerson10.Text;
            txtAuthPerson10.Text = "";


            txtAuthPersonsex9.CheckState = txtAuthPersonsex10.CheckState;
            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }

        private void pictureBox22_Click(object sender, EventArgs e)
        {

            txtAuthPerson10.Text = "";


            txtAuthPersonsex10.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 41;
            AuthPersonlCount -= 1;
        }


        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btnSizeSpecial_Click(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtAuthPersonsex10_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
