using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

using Pop3ClinetApp.Resources;
using LumiSoft.Net.Log;
using LumiSoft.Net.MIME;
using LumiSoft.Net.Mail;
using LumiSoft.Net.POP3.Client;

namespace Pop3ClinetApp
{
    /// <summary>
    /// Application main window.
    /// </summary>
    public class wfrm_Main : Form
    {
        private TabControl m_pTab = null;
        // TabPage Mail
        private ToolStrip m_pTabMail_MessagesToolbar = null;
        private ListView  m_pTabMail_Messages        = null;
        private ListView  m_pTabMail_Attachments     = null;
        private TextBox   m_pTabMail_BodyText        = null;
        // TabPage Log
        private RichTextBox m_pTabLog_LogText = null;

        private POP3_Client m_pPop3 = null;

        /// <summary>
        /// Default constructor.
        /// </summary>
        public wfrm_Main()
        {
            InitUI();

            this.Visible = true;

            wfrm_Connect frm = new wfrm_Connect(new EventHandler<WriteLogEventArgs>(Pop3_WriteLog));
            if(frm.ShowDialog(this) == DialogResult.OK){
                m_pPop3 = frm.POP3;                

                FillMessagesList();
            }
            else{
                Dispose();
            }
        }
                
        #region method Dispose

        /// <summary>
        /// Cleans up any resources being used.
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            // Clean up POP3 client.
            if(m_pPop3 != null){
                m_pPop3.Dispose();
                m_pPop3 = null;
            }
        }

        #endregion

        #region method InitUI

        /// <summary>
        /// Creates and initializes UI.
        /// </summary>
        private void InitUI()
        {
            this.ClientSize = new Size(700,500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Icon = ResManager.GetIcon("app.ico");
            this.Text = "POP3 Client demo";

            m_pTab = new TabControl();
            m_pTab.Dock = DockStyle.Fill;

            #region TabPage Mail

            m_pTab.TabPages.Add("Mail");
            m_pTab.TabPages[0].ClientSize = new Size(700,500); 

            m_pTabMail_MessagesToolbar = new ToolStrip();            
            m_pTabMail_MessagesToolbar.Dock = DockStyle.None;
            m_pTabMail_MessagesToolbar.Location = new Point(595,5);
            m_pTabMail_MessagesToolbar.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            m_pTabMail_MessagesToolbar.GripStyle = ToolStripGripStyle.Hidden;
            m_pTabMail_MessagesToolbar.BackColor = this.BackColor;
            m_pTabMail_MessagesToolbar.Renderer = new ToolBarRendererEx();
            m_pTabMail_MessagesToolbar.ItemClicked += new ToolStripItemClickedEventHandler(m_pTabMail_MessagesToolbar_ItemClicked);
            // Save button
            ToolStripButton button_Save = new ToolStripButton();
            button_Save.Enabled = false;
            button_Save.Image = ResManager.GetIcon("save.ico").ToBitmap();
            button_Save.Name = "save";
            button_Save.Tag = "save";
            button_Save.ToolTipText  = "Save";
            m_pTabMail_MessagesToolbar.Items.Add(button_Save);
            // Delete button
            ToolStripButton button_Delete = new ToolStripButton();
            button_Delete.Enabled = false;
            button_Delete.Image = ResManager.GetIcon("delete.ico").ToBitmap();
            button_Delete.Name = "delete";
            button_Delete.Tag = "delete";
            button_Delete.ToolTipText  = "Delete";
            m_pTabMail_MessagesToolbar.Items.Add(button_Delete);

            m_pTabMail_Messages = new ListView();
            m_pTabMail_Messages.Size = new Size(690,200);
            m_pTabMail_Messages.Location = new Point(5,30);
            m_pTabMail_Messages.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabMail_Messages.View = View.Details;
            m_pTabMail_Messages.HideSelection = false;
            m_pTabMail_Messages.FullRowSelect = true;
            m_pTabMail_Messages.Columns.Add("From",100);
            m_pTabMail_Messages.Columns.Add("Subject",300);
            m_pTabMail_Messages.Columns.Add("Received",120);
            m_pTabMail_Messages.Columns.Add("Size",60);
            m_pTabMail_Messages.SelectedIndexChanged += new EventHandler(m_pTabMail_Messages_SelectedIndexChanged);

            m_pTabMail_Attachments = new ListView();
            m_pTabMail_Attachments.Size = new Size(690,40);
            m_pTabMail_Attachments.Location = new Point(5,240);
            m_pTabMail_Attachments.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabMail_Attachments.View = View.SmallIcon;
            m_pTabMail_Attachments.MouseClick += new MouseEventHandler(m_pAttachments_MouseClick);

            m_pTabMail_BodyText = new TextBox();
            m_pTabMail_BodyText.Size = new Size(690,200);
            m_pTabMail_BodyText.Location = new Point(5,285);
            m_pTabMail_BodyText.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            m_pTabMail_BodyText.ScrollBars = ScrollBars.Both;
            m_pTabMail_BodyText.Multiline = true;

            m_pTab.TabPages[0].Controls.Add(m_pTabMail_MessagesToolbar);
            m_pTab.TabPages[0].Controls.Add(m_pTabMail_Messages);
            m_pTab.TabPages[0].Controls.Add(m_pTabMail_Attachments);
            m_pTab.TabPages[0].Controls.Add(m_pTabMail_BodyText);

            #endregion

            #region TabPage Log

            m_pTab.TabPages.Add("Log");
            m_pTab.TabPages[1].ClientSize = new Size(700,500); 

            m_pTabLog_LogText = new RichTextBox();
            m_pTabLog_LogText.Dock = DockStyle.Fill;
            m_pTabLog_LogText.ReadOnly = true;

            m_pTab.TabPages[1].Controls.Add(m_pTabLog_LogText);

            #endregion

            this.Controls.Add(m_pTab);
        }
                                                
        #endregion


        #region Events Handling

        #region method m_pTabMail_MessagesToolbar_ItemClicked

        private void m_pTabMail_MessagesToolbar_ItemClicked(object sender,ToolStripItemClickedEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            try{
                if(string.Equals(e.ClickedItem.Name,"save")){
                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.FileName = "message.eml";
                    if(dlg.ShowDialog() == DialogResult.OK){
                        this.Cursor = Cursors.WaitCursor;
                        POP3_ClientMessage message = (POP3_ClientMessage)m_pTabMail_Messages.SelectedItems[0].Tag;
                        File.WriteAllBytes(dlg.FileName,message.MessageToByte());
                        this.Cursor = Cursors.Default;
                    }
                }
                else if(string.Equals(e.ClickedItem.Name,"delete")){
                    if(MessageBox.Show(this,"Do you want to delete selected message ?","Confirm Delete:",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes){
                        POP3_ClientMessage message = (POP3_ClientMessage)m_pTabMail_Messages.SelectedItems[0].Tag;
                        message.MarkForDeletion();
                        m_pTabMail_Messages.SelectedItems[0].Remove();
                    }
                }
            }
            catch(Exception x){
                MessageBox.Show(this,"Error: " + x.Message,"Error:",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region method m_pTabMail_Messages_SelectedIndexChanged

        private void m_pTabMail_Messages_SelectedIndexChanged(object sender,EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            try{
                m_pTabMail_MessagesToolbar.Items["save"].Enabled = false;
                m_pTabMail_MessagesToolbar.Items["delete"].Enabled = false;
                if(m_pTabMail_Messages.SelectedItems.Count > 0){
                    m_pTabMail_Attachments.Items.Clear();
                    m_pTabMail_BodyText.Text = "";

                    POP3_ClientMessage message = (POP3_ClientMessage)m_pTabMail_Messages.SelectedItems[0].Tag;
                    Mail_Message       mime    = Mail_Message.ParseFromByte(message.MessageToByte());

                    foreach(MIME_Entity entity in mime.Attachments){
                        ListViewItem item = new ListViewItem();
                        if(entity.ContentDisposition != null && entity.ContentDisposition.Param_FileName != null){
                            item.Text = entity.ContentDisposition.Param_FileName;
                        }
                        else{
                            item.Text = "untitled";
                        }
                        item.Tag = entity;
                        m_pTabMail_Attachments.Items.Add(item);
                    }

                    if(mime.BodyText != null){
                        m_pTabMail_BodyText.Text = mime.BodyText;
                    }

                    m_pTabMail_MessagesToolbar.Items["save"].Enabled = true;
                    m_pTabMail_MessagesToolbar.Items["delete"].Enabled = true;
                }
            }
            catch(Exception x){
                MessageBox.Show(this,"Error: " + x.Message,"Error:",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region method m_pAttachments_MouseClick

        private void m_pAttachments_MouseClick(object sender,MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right && m_pTabMail_Attachments.SelectedItems.Count > 0){
                ContextMenuStrip menu = new ContextMenuStrip();
                menu.Items.Add("Save");
                menu.ItemClicked += new ToolStripItemClickedEventHandler(menu_ItemClicked);
                menu.Show(Control.MousePosition);                
            }
        }

        private void menu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            MIME_Entity entity = (MIME_Entity)m_pTabMail_Attachments.SelectedItems[0].Tag;
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = m_pTabMail_Attachments.SelectedItems[0].Text;
            if(dlg.ShowDialog(this) == DialogResult.OK){
                File.WriteAllBytes(dlg.FileName,((MIME_b_SinglepartBase)entity.Body).Data);
            }
        }

        #endregion


        #region method Pop3_WriteLog

        private delegate void AppendText(string text);

        /// <summary>
        /// This method is called when POP3 client has new log entry.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event data.</param>
        private void Pop3_WriteLog(object sender,WriteLogEventArgs e)
        {
            try{
                if(e.LogEntry.EntryType == LogEntryType.Read){
                    m_pTabLog_LogText.BeginInvoke(new AppendText(m_pTabLog_LogText.AppendText),ObjectToString(e.LogEntry.RemoteEndPoint) + " >> " + e.LogEntry.Text + "\r\n");
                }
                else if(e.LogEntry.EntryType == LogEntryType.Write){
                    m_pTabLog_LogText.BeginInvoke(new AppendText(m_pTabLog_LogText.AppendText),ObjectToString(e.LogEntry.RemoteEndPoint) + " << " + e.LogEntry.Text + "\r\n");
                }
                else if(e.LogEntry.EntryType == LogEntryType.Text){
                    m_pTabLog_LogText.BeginInvoke(new AppendText(m_pTabLog_LogText.AppendText),ObjectToString(e.LogEntry.RemoteEndPoint) + " xx " + e.LogEntry.Text + "\r\n");
                }
            }
            catch(Exception x){
                MessageBox.Show(x.ToString());
            }
        }

        #endregion

        #endregion


        #region method FillMessagesList

        /// <summary>
        /// Gets messages list from POP3 server and adds them to UI.
        /// </summary>
        private void FillMessagesList()
        {
            this.Cursor = Cursors.WaitCursor;
            try{
                foreach(POP3_ClientMessage message in m_pPop3.Messages){
                    Mail_Message mime = Mail_Message.ParseFromByte(message.HeaderToByte());

                    ListViewItem item = new ListViewItem();
                    if(mime.From != null){
                        item.Text = mime.From.ToString();
                    }
                    else{
                        item.Text = "<none>";
                    }
                    if(string.IsNullOrEmpty(mime.Subject)){
                        item.SubItems.Add("<none>");
                    }
                    else{
                        item.SubItems.Add(mime.Subject);
                    }
                    item.SubItems.Add(mime.Date.ToString());
                    item.SubItems.Add(((decimal)(message.Size / (decimal)1000)).ToString("f2") + " kb");
                    item.Tag = message;
                    m_pTabMail_Messages.Items.Add(item);
                }
            }
            catch(Exception x){
                MessageBox.Show(this,"Error: " + x.Message,"Error:",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            this.Cursor = Cursors.Default;
        }

        #endregion

        #region method ObjectToString

        /// <summary>
        /// Calls obj.ToSting() if obj is not null, otherwise returns "".
        /// </summary>
        /// <param name="obj">Object.</param>
        /// <returns>Returns obj.ToSting() if obj is not null, otherwise returns "".</returns>
        private string ObjectToString(object obj)
        {
            if(obj == null){
                return "";
            }
            else{
                return obj.ToString();
            }
        }

        #endregion

    }
}
