using LSEXT;
using LSSERVICEPROVIDERLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RightClick_ARM
{
    [ComVisible(true)]
    [ProgId("RightClick_ARM.RightClickArm")]
    public class RightClickArm : IEntityExtension
    {
        INautilusServiceProvider sp;

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }
        ContainerFrm frm;
        public void Execute(ref LSExtensionParameters Parameters)
        {
            //string tableName = Parameters["TABLE_NAME"];

            sp = Parameters["SERVICE_PROVIDER"];
            var rs = Parameters["RECORDS"];

            string sampleMsgName = rs[0].Value;

            frm = new ContainerFrm(sampleMsgName);
            AssutaRequestManager.AssutaRequestManagerCls arm = new AssutaRequestManager.AssutaRequestManagerCls();

            arm.RunByEntityExtention(sampleMsgName,sp);
            arm.Dock = DockStyle.Fill;
            frm.Controls.Add(arm);

            
            AssutaRequestManager.AssutaRequestManagerCls.Closedddd += arm_Closedddd;
            frm.WindowState = FormWindowState.Maximized;
            frm.ShowDialog();
            frm.BringToFront();
            AssutaRequestManager.AssutaRequestManagerCls.Closedddd -= arm_Closedddd;
        }

        void arm_Closedddd(bool isChanged)
        {
            if (isChanged)
            {
                DialogResult result = System.Windows.Forms.MessageBox.Show("ישנם נתונים שלא נשמרו, האם ברצונך לצאת?", "יציאה מהחלון", MessageBoxButtons.OKCancel);
                if (result == DialogResult.OK)
                {
                    frm.Close();
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                frm.Close();
            }
        }
    }
}
