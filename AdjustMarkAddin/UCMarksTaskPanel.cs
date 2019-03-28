using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace AdjustMarkAddin
{
    public partial class UCMarksTaskPanel : UserControl
    {
        public UCMarksTaskPanel()
        {
            InitializeComponent();
            InitDefaultValue();
        }

        private void InitDefaultValue()
        {
            cmbACSSubType.SelectedIndex = cmbACSSubType.Items.IndexOf("幅值");
            cmbACSValueType.SelectedIndex = cmbACSValueType.Items.IndexOf("误差值");
            cmbACSPhase.SelectedIndex = cmbACSPhase.Items.IndexOf("A");
            cmbACSTestModeP.SelectedIndex = cmbACSTestModeP.Items.IndexOf("4L");
            cmbACSPowerTypeP.SelectedIndex = cmbACSPowerTypeP.Items.IndexOf("有功");
            cmbACSValueTypeP.SelectedIndex = cmbACSValueTypeP.Items.IndexOf("误差值");
            cmbACSPhaseP.SelectedIndex = cmbACSPhaseP.Items.IndexOf("总");
            cmbACSValueTypeF.SelectedIndex = cmbACSValueTypeF.Items.IndexOf("误差值");
            cmbACSPhaseS.SelectedIndex = cmbACSPhaseS.Items.IndexOf("Ua");


            cmbDCSValueType.SelectedIndex = cmbDCSValueType.Items.IndexOf("误差值");
            cmbDCMDirection.SelectedIndex = cmbDCMDirection.Items.IndexOf("正向");
            cmbDCMValueType.SelectedIndex = cmbDCMValueType.Items.IndexOf("误差值");
        }

        private void GetACSTypes(out string subType, out string valueType, out string phaseType)
        {
            subType = string.Empty;
            switch (cmbACSSubType.SelectedIndex)
            {
                case 0:
                    subType = "amplitude";
                    break;
                case 1:
                    subType = "phase";
                    break;
                default:
                    subType = "amplitude";
                    break;
            }
            valueType = string.Empty;
            switch (cmbACSValueType.SelectedIndex)
            {
                case 0:
                    valueType = "err";
                    break;
                case 1:
                    valueType = "ref";
                    break;
                case 2:
                    valueType = "mea";
                    break;
                case 3:
                    valueType = "con";
                    break;
                default:
                    valueType = "err";
                    break;
            }
            phaseType = string.Empty;
            switch (cmbACSPhase.SelectedIndex)
            {
                case 0:
                    phaseType = "A";
                    break;
                case 1:
                    phaseType = "B";
                    break;
                case 2:
                    phaseType = "C";
                    break;
                default:
                    phaseType = "A";
                    break;
            }
        }

        private void btnAddACSV_Click(object sender, EventArgs e)
        {
            GetACSTypes(out string subType, out string valueType, out string phaseType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("ACSourceVoltage;{0}V_{1}%;{2};{3};{4}", cmbACSVRange.Text, cmbACSVPercent.Text, subType, valueType, phaseType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddACSC_Click(object sender, EventArgs e)
        {
            GetACSTypes(out string subType, out string valueType, out string phaseType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("ACSourceCurrent;{0}A_{1}%;{2};{3};{4}", cmbACSCRange.Text, cmbACSCPercent.Text, subType, valueType, phaseType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void GetACSTypesP(out string testMode, out string subType, out string valueType, out string phaseType)
        {
            testMode = string.Empty;
            switch (cmbACSTestModeP.SelectedIndex)
            {
                case 0:
                    testMode = "4L";
                    break;
                case 1:
                    testMode = "3L";
                    break;
                default:
                    testMode = "4L";
                    break;
            }
            subType = string.Empty;
            switch (cmbACSPowerTypeP.SelectedIndex)
            {
                case 0:
                    subType = "active";
                    break;
                case 1:
                    subType = "reactive";
                    break;
                default:
                    subType = "active";
                    break;
            }
            valueType = string.Empty;
            switch (cmbACSValueTypeP.SelectedIndex)
            {
                case 0:
                    valueType = "err";
                    break;
                case 1:
                    valueType = "ref";
                    break;
                case 2:
                    valueType = "mea";
                    break;
                case 3:
                    valueType = "con";
                    break;
                default:
                    valueType = "err";
                    break;
            }
            phaseType = string.Empty;
            switch (cmbACSPhaseP.SelectedIndex)
            {
                case 0:
                    phaseType = "T";
                    break;
                case 1:
                    phaseType = "A";
                    break;
                case 2:
                    phaseType = "B";
                    break;
                case 3:
                    phaseType = "C";
                    break;
                default:
                    phaseType = "T";
                    break;
            }
        }

        private void btnAddACSP_Click(object sender, EventArgs e)
        {
            GetACSTypesP(out string testMode, out string subType, out string valueType, out string phaseType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("ACSourcePower;{0}_{1}V_{2}%_{3}A_{4}%_{5};{6};{7};{8}", testMode, cmbACSVRangeP.Text, cmbACSVPercentP.Text, cmbACSCRangeP.Text, cmbACSCPercentP.Text, cmbACSFactorP.Text, subType, valueType, phaseType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddACSF_Click(object sender, EventArgs e)
        {
            string valueType = string.Empty;
            switch (cmbACSValueTypeP.SelectedIndex)
            {
                case 0:
                    valueType = "err";
                    break;
                case 1:
                    valueType = "ref";
                    break;
                case 2:
                    valueType = "mea";
                    break;
                case 3:
                    valueType = "con";
                    break;
                default:
                    valueType = "err";
                    break;
            }
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("ACSourceFrequency;{0}Hz_{1}V_{2}%;frequency;{3};T", cmbACSFrequencyF.Text, cmbACSVRangeF.Text, cmbACSVPercentF.Text, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddACSS_Click(object sender, EventArgs e)
        {
            string phaseType = string.Empty;
            switch (cmbACSPhaseS.SelectedIndex)
            {
                case 0:
                    phaseType = "Ua";
                    break;
                case 1:
                    phaseType = "Ub";
                    break;
                case 2:
                    phaseType = "Uc";
                    break;
                case 3:
                    phaseType = "Ia";
                    break;
                case 4:
                    phaseType = "Ib";
                    break;
                case 5:
                    phaseType = "Ic";
                    break;
                case 6:
                    phaseType = "Pa";
                    break;
                case 7:
                    phaseType = "Pb";
                    break;
                case 8:
                    phaseType = "Pc";
                    break;
                case 9:
                    phaseType = "T";
                    break;
                default:
                    phaseType = "T";
                    break;
            }
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("ACSourceStability;{0}V_{1}%_{2}A_{3}%_{4};stability;sta;{5}", cmbACSVRangeS.Text, cmbACSVPercentS.Text, cmbACSCRangeS.Text, cmbACSCPercentS.Text, cmbACSFactorS.Text, phaseType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }


        private void GetDCSTypes(out string valueType)
        {
            valueType = string.Empty;
            switch (cmbDCSValueType.SelectedIndex)
            {
                case 0:
                    valueType = "err";
                    break;
                case 1:
                    valueType = "ref";
                    break;
                case 2:
                    valueType = "mea";
                    break;
                case 3:
                    valueType = "con";
                    break;
                default:
                    valueType = "err";
                    break;
            }
        }

        private void btnAddDCSV_Click(object sender, EventArgs e)
        {
            GetDCSTypes(out string valueType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("DCSourceVoltage;{0}V_{1}%;amplitude;{2};T", cmbDCSVRange.Text, cmbDCSVPercent.Text, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddDCSC_Click(object sender, EventArgs e)
        {
            GetDCSTypes(out string valueType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("DCSourceCurrent;{0}A_{1}%;amplitude;{2};T", cmbDCSCRange.Text, cmbDCSCPercent.Text, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void GetDCMTypes(out string direction, out string valueType)
        {
            direction = string.Empty;
            switch (cmbDCMDirection.SelectedIndex)
            {
                case 0:
                    direction = "forward";
                    break;
                case 1:
                    direction = "reverse";
                    break;
                default:
                    direction = "forward";
                    break;
            }
            valueType = string.Empty;
            switch (cmbDCMValueType.SelectedIndex)
            {
                case 0:
                    valueType = "err";
                    break;
                case 1:
                    valueType = "ref";
                    break;
                case 2:
                    valueType = "mea";
                    break;
                case 3:
                    valueType = "con";
                    break;
                default:
                    valueType = "err";
                    break;
            }
        }

        private void btnAddDCMV_Click(object sender, EventArgs e)
        {
            GetDCMTypes(out string direction, out string valueType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("DCMeterVoltage;{0}V_{1}%;{2};{3};T", cmbDCMVRange.Text, cmbDCMVPercent.Text, direction, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddDCMVS_Click(object sender, EventArgs e)
        {
            GetDCMTypes(out string direction, out string valueType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("DCMeterVoltage2;{0}V_{1}%;{2};{3};T", cmbDCMVSRange.Text, cmbDCMVSPercent.Text, direction, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }

        private void btnAddDCMC_Click(object sender, EventArgs e)
        {
            GetDCMTypes(out string direction, out string valueType);
            object range = Globals.ThisAddIn.Application.Selection.Range;
            ((Microsoft.Office.Interop.Word.Range)range).Text = string.Format("DCMeterCurrent;{0}A_{1}%;{2};{3};T", cmbDCMCRange.Text, cmbDCMCPercent.Text, direction, valueType);
            Bookmark bookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(string.Format("A{0:yyyyMMddHHmmssfff}", DateTime.Now), ref range);
            bookmark.Select();
        }
    }
}
