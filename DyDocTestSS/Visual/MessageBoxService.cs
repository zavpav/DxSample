using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Services;

namespace DyDocTestSS.Visual
{
    public class MessageBoxService : IMessageBoxService
    {
        public IMessageBoxService MsgSrv { get; }

        public MessageBoxService(IMessageBoxService msgSrv)
        {
            MsgSrv = msgSrv;
        }

        public DialogResult ShowMessage(string message, string title, MessageBoxIcon icon)
        {
            if (message.Contains("Снять защиту листа"))
                return DialogResult.OK;

            return this.MsgSrv.ShowMessage(message, title, icon);
        }

        public DialogResult ShowDataValidationDialog(string message, string title, DataValidationErrorStyle errorStyle)
        {
            return this.MsgSrv.ShowDataValidationDialog(message, title, errorStyle);
        }

        public DialogResult ShowYesNoCancelMessage(string message)
        {
            return this.MsgSrv.ShowYesNoCancelMessage(message);
        }

        public bool ShowOkCancelMessage(string message)
        {
            return this.MsgSrv.ShowOkCancelMessage(message);
        }

        public bool ShowYesNoMessage(string message)
        {
            return this.MsgSrv.ShowYesNoMessage(message);
        }
    }
}