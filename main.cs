using System;
using System.Drawing;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
 
namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
        }
 
        private void Application_DocumentOpen(Document Doc)
        {
            Doc.ContentControlOnEnter += new DocumentContentControlEvents_ContentControlOnEnterEventHandler(Doc_ContentControlOnEnter);
        }
 
        private void Doc_ContentControlOnEnter(ContentControl ContentControl)
        {
            CommandBar contextMenu = this.Application.CommandBars["Text"];
            CommandBarButton underlineButton = (CommandBarButton)contextMenu.Controls.Add(
                MsoControlType.msoControlButton, missing, missing, missing, true);
            underlineButton.Caption = "Custom Underline";
            underlineButton.Click += new _CommandBarButtonEvents_ClickEventHandler(UnderlineButton_Click);
        }
 
        private void UnderlineButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            string textToUnderline = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter the text to underline:", "Custom Underline");
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.ShowDialog();
            Color underlineColor = colorDialog.Color;
 
            Selection selection = this.Application.Selection;
            selection.HomeKey(WdUnits.wdStory, WdMovementType.wdMove);
 
            while (selection.Find.Execute(textToUnderline))
            {
                selection.Range.Underline = WdUnderline.wdUnderlineWavy;
                selection.Range.Font.UnderlineColor = (WdColor)(int)underlineColor.ToArgb() & 0xFFFFFF;
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
        }
 
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }
 
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}
 

//Please donote us!!
//BTC - bc1q5kmqqynratseyh7v0n8q58rn7p5xejuemmc4px
//USDT(ETH) - 0x8558288490E11E7F900471E7D52F0b0A0B6b8572
//USDT(SOLANA) - 4MjmiAwiQT1cqb5fSpvdsKCabZAKxopcMsTqem9gWBqB
//USDT(POLYGON) - 0x8558288490E11E7F900471E7D52F0b0A0B6b8572
//ETH - 0x8558288490E11E7F900471E7D52F0b0A0B6b8572
