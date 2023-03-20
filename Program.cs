using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(".NET Graph Tutorial\n");

            var settings = Settings.LoadSettings();

            // Initialize Graph
            InitializeGraph(settings);

            // Greet the user by name
            GreetUserAsync();

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Display access token");
                Console.WriteLine("2. List my inbox");
                Console.WriteLine("3. Send mail");
                Console.WriteLine("4. List users (requires app-only)");
                Console.WriteLine("5. Make a Graph call");

                try
                {
                    choice = int.Parse(Console.ReadLine() ?? string.Empty);
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch (choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;
                    case 1:
                        // Display access token
                        DisplayAccessTokenAsync();
                        break;
                    case 2:
                        // List emails from user's inbox
                        ListInboxAsync();
                        break;
                    case 3:
                        // Send an email message
                        SendMailAsync();
                        break;
                    //case 4:
                    //    // List users
                    //    await ListUsersAsync();
                    //    break;
                    //case 5:
                    //    // Run any Graph code
                    //    await MakeGraphCallAsync();
                    //    break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }


        static void InitializeGraph(Settings settings)
        {
            GraphHelper.EnsureGraphForAppOnlyAuth(settings);
            return;
            GraphHelper.InitializeGraphForUserAuth(settings,
                (info, cancel) =>
                {
                    // Display the device code message to
                    // the user. This tells them
                    // where to go to sign in and provides the
                    // code to use.
                    Console.WriteLine(info.Message);
                    return Task.FromResult(0);
                });
        }

        static async Task DisplayAccessTokenAsync()
        {
            try
            {
                var userToken = await GraphHelper.GetUserTokenAsync();
                Console.WriteLine($"User token: {userToken}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting user access token: {ex.Message}");
            }
        }


        static async Task GreetUserAsync()
        {
            // TODO
        }

        static async Task SendMailAsync()
        {
            try
            {
                // Send mail to the signed-in user
                // Get the user for their email address
                //var user = await GraphHelper.GetUserAsync();

                //  var userEmail = "ITC-Msg.SharedMbx08@Vodafone-itc.com"; 
                var userEmail = "markos.kitrinakis@gmail.com";
                if (string.IsNullOrEmpty(userEmail))
                {
                    Console.WriteLine("Couldn't get your email address, canceling...");
                    return;
                }

                //await GraphHelper.SendMailAsync("Testing Microsoft Graph","Hello world!", userEmail);
                GraphHelper.SendMailAsync("Testing Microsoft Graph", "Hello world!", userEmail);
                Console.WriteLine("Mail sent.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending mail: {ex.Message}");
            }
        }

        static void ListInboxAsync()
        {
            try
            {
                var messagePage = GraphHelper.GetInbox();

                // Output each message's details
                foreach (var message in messagePage.CurrentPage)
                {
                    Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
                    Console.WriteLine($"  From: {message.From?.EmailAddress?.Name}");

                    Console.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
                }

                // If NextPageRequest is not null, there are more messages
                // available on the server
                // Access the next page like:
                // messagePage.NextPageRequest.GetAsync();
                var moreAvailable = messagePage.NextPageRequest != null;

                Console.WriteLine($"\nMore messages available? {moreAvailable}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting user's inbox: {ex.Message}");
            }
        }
    }
}
