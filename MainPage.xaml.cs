using Aspose.Email;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Windows.Storage;
using Windows.Storage.AccessCache;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace AsposeNTC
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private async void SetLicense_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FileOpenPicker openPicker = new FileOpenPicker
                {
                    ViewMode = PickerViewMode.List,
                    SuggestedStartLocation = PickerLocationId.DocumentsLibrary
                };
                openPicker.FileTypeFilter.Add(".lic");
                var file = await openPicker.PickSingleFileAsync();
                if (file != null)
                {
                    var token = StorageApplicationPermissions.FutureAccessList.Add(file);
                    StorageFolder localFolder = ApplicationData.Current.LocalFolder;
                    await file.CopyAsync(localFolder, file.Name, NameCollisionOption.ReplaceExisting);
                    // Read the license file from local folder
                    byte[] licenseBytes = File.ReadAllBytes(Path.Combine(localFolder.Path, file.Name));

                    // Create memory stream from the byte array
                    using (MemoryStream licenseStream = new MemoryStream(licenseBytes))
                    {
                        // Create Aspose.Email License instance
                        License license = new License();

                        // Set the license using the memory stream
                        license.SetLicense(licenseStream);
                        TB1.Text = "License set successfully.";
                    }
                }
            }
            catch (Exception ex)
            {
                TB1.Text = ex.ToString() + ex.InnerException?.ToString();
            }
        }

        private async void ReadPST_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FileOpenPicker openPicker = new FileOpenPicker
                {
                    ViewMode = PickerViewMode.List,
                    SuggestedStartLocation = PickerLocationId.DocumentsLibrary
                };
                openPicker.FileTypeFilter.Add(".pst");
                var file = await openPicker.PickSingleFileAsync();
                if (file != null)
                {
                    var token = StorageApplicationPermissions.FutureAccessList.Add(file);
                    StorageFolder localFolder = ApplicationData.Current.LocalFolder;
                    await file.CopyAsync(localFolder, file.Name, NameCollisionOption.ReplaceExisting);

                    PersonalStorage personalStorage = PersonalStorage.FromFile(Path.Combine(localFolder.Path, file.Name));
                    if (personalStorage != null)
                    {
                        TB2.Text = "Folder count : " + personalStorage.RootFolder.GetSubFolders().Count.ToString();
                        var folderInfo = personalStorage.RootFolder.GetSubFolder("Inbox");
                        if (folderInfo != null)
                        {
                            var messageInfoCollection = folderInfo.GetContents();
                            List<string> subjects = new List<string>();
                            foreach (var messageInfo in messageInfoCollection)
                            {
                                subjects.Add(messageInfo.Subject);
                            }
                            TB2.Text = TB2.Text + '\n' + string.Join(Environment.NewLine, subjects);
                        }
                    }
                    else
                    {
                        TB2.Text = "PersonalStorage is null";
                    }
                }
            }
            catch (Exception ex)
            {
                TB2.Text = ex.ToString() + ex.InnerException?.ToString();
            }
        }

        private async void ReadEML_Click(object sender, RoutedEventArgs e)
        { 
            try
            {
                FileOpenPicker openPicker = new FileOpenPicker
                {
                    ViewMode = PickerViewMode.List,
                    SuggestedStartLocation = PickerLocationId.DocumentsLibrary
                };
                openPicker.FileTypeFilter.Add(".eml");
                var file = await openPicker.PickSingleFileAsync();
                if (file != null)
                {
                    var token = StorageApplicationPermissions.FutureAccessList.Add(file);
                    StorageFolder localFolder = ApplicationData.Current.LocalFolder;
                    await file.CopyAsync(localFolder, file.Name, NameCollisionOption.ReplaceExisting);

                    StorageFile storagefile = await localFolder.GetFileAsync(file.Name);

                    string fileContents = await FileIO.ReadTextAsync(storagefile);
                    byte[] byteArray = Encoding.UTF8.GetBytes(fileContents);

                    MemoryStream fileStream = new MemoryStream(byteArray);

                    MailMessage message = MailMessage.Load(fileStream);

                    if (message != null)
                    {
                        TB3.Text = "Mail Subject : " + message.Subject;
                    }
                    else
                    {

                        TB3.Text = "MailMessage is null";
                    }
                }
            }
            catch (Exception ex)
            {
                TB3.Text = ex.ToString() + ex.InnerException?.ToString();
            }
        }
    }
}
