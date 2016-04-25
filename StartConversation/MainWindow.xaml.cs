using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Documents;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Extensibility;
using MessageBox = System.Windows.MessageBox;
using System.Linq;


namespace StartConversation
{
    public partial class MainWindow : Window
    {
        Microsoft.Lync.Model.LyncClient client = null;
        Microsoft.Lync.Model.Extensibility.Automation automation = null;
        List<Contact> ccts = new List<Contact>();

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                //Start the conversation
                automation = LyncClient.GetAutomation();
                client = LyncClient.GetClient();
                ClientState status = client.State;
                if (client.State != ClientState.SignedIn)
                {
                    throw new Exception("Client is not Signed in. Launch Lync/Skype for Business client e reopen the application.");
                }
            }
            catch (LyncClientException )
            {
                txtErrors.Text = "Error: Failed to connect to Lync.";
            }
            catch (Exception err)
            {
                // Rethrow the SystemException which did not come from the Lync Model API.
                txtErrors.Text = "Error: " + err.Message;
            }
        }

        private void btnStartConv_Click(object sender, RoutedEventArgs e)
        {
            //Clean Richbox text and Contacts List
            rtbParticipants.Document.Blocks.Clear();
            ccts.Clear();
            GetOnlineContacts();
        }

        private static char[] Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        private void FillContactsList()
        {
            List<SkypeContacts> listContacts = new List<SkypeContacts>();
            string presenceStatus;
            int contactType;
            int totalContacts = 0;

            foreach (Contact obj in ccts)
            {
                if (ccts.Count == 0)
                {
                    MessageBox.Show("Error: 0 contacts found!");
                }
                else
                {
                    presenceStatus = obj.GetContactInformation(ContactInformationType.Activity).ToString();

                    if (isValidPresence(presenceStatus.ToLower()))
                    {
                        contactType = Convert.ToInt32(obj.GetContactInformation(ContactInformationType.SourceNetwork));
                        //Contact Status 0 and 1 are related for internal Lync Users
                        //ref: https://msdn.microsoft.com/en-us/library/microsoft.lync.model.sourcenetworktype_di_3_uc_ocs14mreflyncclnt(v=office.14).aspx
                        if (contactType == 0 || contactType == 1)
                            listContacts.Add(new SkypeContacts(obj.Uri, presenceStatus));
                    }
                }
            }
            totalContacts = ccts.Count;
            List<SkypeContacts> notDuplicatedContacts = RemoveDuplicatedItens(listContacts);

            //This part of code call the method AddContactsToField at the same thread
            Dispatcher.Invoke(new Action(() => { txtErrors.Text = string.Format("{0}: Total Contacts Found | {1}: Online Contacts", totalContacts, notDuplicatedContacts.Count); }), System.Windows.Threading.DispatcherPriority.ContextIdle);
            Dispatcher.Invoke(new Action(() => { AddContactsToField(notDuplicatedContacts); }), System.Windows.Threading.DispatcherPriority.ContextIdle);

        }

        private void AddContactsToField(List<SkypeContacts> contacts)
        {
            //It require to add a new line before all the append texts to
            //not break the first line
            rtbParticipants.AppendText(System.Environment.NewLine);
            foreach (SkypeContacts obj in contacts)
            {
                if (!string.IsNullOrEmpty(obj.Sip))
                    rtbParticipants.AppendText(obj.Sip.Trim() + System.Environment.NewLine);
            }
        }

        private List<SkypeContacts> RemoveDuplicatedItens(List<SkypeContacts> list)
        {
            return list.Distinct(new ContactEqualityComparer()).ToList();
        }

        private bool isValidPresence(string presence)
        {
            bool isValid = false;

            //Pick only user who are avaiable to receive messages
            if (presence == "available" || presence == "away" || presence == "busy" || presence == "inactive" || presence == "in a meeting")
            {
                isValid = true;

            }
            return isValid;
        }

        private void GetOnlineContacts()
        {
            SearchFields searchFields = client.ContactManager.GetSearchFields();
            SearchProviders chosenProviders = SearchProviders.GlobalAddressList;
            int initialLetterIndex = 0;
            try
            {
                IAsyncResult result = client.ContactManager.BeginSearch(Alphabet[initialLetterIndex].ToString(), chosenProviders, searchFields, SearchOptions.Default, 5000, SearchAllCallback, new object[] { initialLetterIndex, new List<Contact>() });

                txtErrors.Text = "Finding contacts in Global catalog...";

                client.ContactManager.EndSearch(result);
            }
            catch (Exception err)
            {
                txtErrors.Text = string.Format("Error: {0}", err.Message);
            }
        }

        private void SearchAllCallback(IAsyncResult result)
        {
            object[] parameters = (object[])result.AsyncState;
            int letterIndex = (int)parameters[0] + 1;
            List<Contact> contacts = (List<Contact>)parameters[1];
            SearchResults results = null;

            if (letterIndex < Alphabet.Length)
            {
                results = client.ContactManager.EndSearch(result);
                contacts.AddRange(results.Contacts);
                client.ContactManager.BeginSearch(Alphabet[letterIndex].ToString(), SearchAllCallback, new object[] { letterIndex, contacts });
            }
            else
            {
                ccts.AddRange(contacts);
                FillContactsList();
            }
        }


        private void btnSendMessage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TextRange tr = new TextRange(rtbParticipants.Document.ContentStart, rtbParticipants.Document.ContentEnd);
                if (String.IsNullOrEmpty(tr.Text.Trim()))
                {
                    txtErrors.Text = "No participants specified!";
                    return;
                }

                //Add all contacts in an array and remove empty entries
                String[] participants = tr.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                int partCount = participants.Count();

                for (int i = 0; i < partCount; i++)
                {
                    TextRange textRange = new TextRange(rtbMessage.Document.ContentStart, rtbMessage.Document.ContentEnd);

                    String broadCastMsg = textRange.Text;

                    if (string.IsNullOrWhiteSpace(broadCastMsg))
                    {
                        txtErrors.Text = "Message field is empty!";
                    }
                    else
                    {
                        SendSyncMessage(broadCastMsg, participants[i]);
                    }
                }
            }
            catch (Exception err)
            {
                txtErrors.Text = string.Format("Error: {0}", err.Message);
            }
        }

        private void SendSyncMessage(string message, string contactUri)
        {
            try
            {
                Contact contact = client.ContactManager.GetContactByUri(contactUri);

                Conversation conversation = client.ConversationManager.AddConversation();
                conversation.AddParticipant(contact);                

                Dictionary<InstantMessageContentType, String> messages = new Dictionary<InstantMessageContentType, String>();
                messages.Add(InstantMessageContentType.PlainText, message);

                InstantMessageModality m = (InstantMessageModality)conversation.Modalities[ModalityTypes.InstantMessage];
                m.BeginSendMessage(messages, null, messages);
            }
            catch (Exception err)
            {
                throw new Exception(err.Message);
            }
        }

        void ConversationTest_IsTypingChanged(object sender, IsTypingChangedEventArgs e)
        {

        }

        void ConversationTest_InstantMessageReceived(object sender, MessageSentEventArgs e)
        {

        }

        private bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }

    }

    public class SkypeContacts
    {
        public string Sip { get; set; }
        public string Presence { get; set; }

        public SkypeContacts(string sip, string presence)
        {
            this.Sip = sip;
            this.Presence = presence;
        }
    }

    class ContactEqualityComparer : IEqualityComparer<SkypeContacts>
    {
        public bool Equals(SkypeContacts x, SkypeContacts y)
        {
            return x.Sip.Equals(y.Sip);
        }

        public int GetHashCode(SkypeContacts obj)
        {
            return obj.Sip.GetHashCode();
        }
    }
}
