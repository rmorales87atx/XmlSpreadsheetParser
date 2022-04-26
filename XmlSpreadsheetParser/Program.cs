using System;
using System.Collections.Generic;

namespace TestApp {
    public class Program {
        static readonly Dictionary<string, string> testValues = new Dictionary<string, string>() {
            { "RPL_WELCOME", "{0} :Welcome to the Internet Relay Network {1}" },
            { "RPL_YOURHOST", "{0} :Your host is {1}, running version {2}" },
            { "RPL_CREATED", "{0} :This server was created {1}" },
            { "RPL_MYINFO", "{0} :{1} {2} {3} {4} {5}" },
            { "RPL_ISUPPORT", "{0} {1} :supported by this server" },
            { "RPL_SNOMASK", "{0} {1} :Server notice mask (0x{2:X})" },
            { "RPL_STATSLINKINFO", "{0} {1} * {2} {3} {4} {5} {6}" },
            { "RPL_STATSCOMMANDS", "{0} {1} {2} {3} {4}" },
            { "RPL_STATSKLINE", "{0} K {1} * {2} {3} :{4}" },
            { "RPL_ENDOFSTATS", "{0} {1} :End of /STATS" },
            { "RPL_STATSWLINE", "{0} W *@{1} * * :{2}" },
            { "RPL_UMODE", "{0} +{1}" },
            { "RPL_STATSUPTIME", "{0} :Server up {1} days {2:00}:{3:00}:{4:00}" },
            { "RPL_STATSOLINE", "{0} O {1} * {2}" },
            { "RPL_LUSERCLIENT", "{0} :{1} user(s), {2} invisible, on {3} server(s)" },
            { "RPL_LUSEROP", "{0} {1} :Operations Staff member(s) available" },
            { "RPL_LUSERUNKNOWN", "{0} {1} :unknown connection(s)" },
            { "RPL_LUSERCHANNELS", "{0} {1} :channels formed" },
            { "RPL_LUSERME", "{0} :I have {1} client(s) and {2} server(s)" },
            { "RPL_ADMINME", "{0} {1} :Administrative information" },
            { "RPL_ADMINLOC1", "{0} :{1}" },
            { "RPL_ADMINLOC2", "{0} :{1}" },
            { "RPL_ADMINEMAIL", "{0} :{1}" },
            { "RPL_WAIT", "{0} {1} :{2}" },
            { "RPL_LOCALUSERS", "{0} {1} {2} :Current local users {1}, max {2}" },
            { "RPL_GLOBALUSERS", "{0} {1} {2} :Current global users {1}, max {2}" },
            { "RPL_USINGSSL", "{0} {1} :is using a secure connection (SSL)" },
            { "RPL_WHOISCERTFP", "{0} {1} :has client certificate fingerprint {2}" },
            { "RPL_JUPELIST", "{0} {1} {2} :{3}" },
            { "RPL_ENDOFJUPELIST", "{0} :End of JUPE list" },
            { "RPL_MISC", "{0} :{1}" },
            { "RPL_ISAWAY", "{0} {1} :{2}" },
            { "RPL_USERHOST", "{0} :{1}" },
            { "RPL_UNAWAY", "{0} :You are no longer marked as being away" },
            { "RPL_AWAY", "{0} :{1}" },
            { "RPL_WHOISHELPER", "{0} {1} :is available for help" },
            { "RPL_WHOISUSER", "{0} {1} {2} {3} * :{4}" },
            { "RPL_WHOISSERVER", "{0} {1} {2} :{3}" },
            { "RPL_WHOISOPERATOR", "{0} {1} :is an IRC operator" },
            { "RPL_ENDWHO", "{0} {1} :End of /WHO" },
            { "RPL_WHOISIDLE", "{0} {1} {2} {3} :seconds idle, signon time" },
            { "RPL_ENDOFWHOIS", "{0} {1} :End of /WHOIS" },
            { "RPL_WHOISCHANNELS", "{0} {1} :{2}" },
            { "RPL_WHOISSVCMSG", "{0} {1} :is an official network service" },
            { "RPL_LISTSTART", "{0} Channel :Users  Topic" },
            { "RPL_LIST", "{0} {1} {2} :{3}" },
            { "RPL_LISTEND", "{0} :End of channel list" },
            { "RPL_CHANNELMODEIS", "{0} {1} {2}" },
            { "RPL_WHOISWEBIRC", "{0} {1} :is connected via {2} ({3})" },
            { "RPL_NOTOPIC", "{0} {1} :No topic is set" },
            { "RPL_TOPIC", "{0} {1} :{2}" },
            { "RPL_TOPICWHOTIME", "{0} {1} {2} {3}" },
            { "RPL_WHOISBOT", "{0} {1} :is a bot" },
            { "RPL_WHOISACTUALLY", "{0} {1} {2}@{3} {4} :Actual user@host, Actual IP" },
            { "RPL_INVITING", "{0} {1} {2}" },
            { "RPL_INVITELIST", "{0} {1} {2}" },
            { "RPL_ENDOFINVITELIST", "{0} {1} :End of channel invite list" },
            { "RPL_EXCEPTLIST", "{0} {1} {2}" },
            { "RPL_ENDOFEXCEPTLIST", "{0} {1} :End of channel exception list" },
            { "RPL_VERSION", "{0} {1} {2} :{3}" },
            { "RPL_WHO", "{0} {1} {2} {3} {4} {5} {6} :0 {7}" },
            { "RPL_NAMES", "{0} {1} {2} :{3}" },
            { "RPL_ENDOFNAMES", "{0} {1} :End of /NAMES" },
            { "RPL_BANLIST", "{0} {1} {2}" },
            { "RPL_ENDOFBANLIST", "{0} {1} :End of channel ban list" },
            { "RPL_INFO", "{0} :- {1}" },
            { "RPL_MOTD", "{0} :- {1}" },
            { "RPL_INFOSTART", "{0} :- {1}" },
            { "RPL_ENDOFINFO", "{0} :{1}" },
            { "RPL_MOTDSTART", "{0} :- Message of the day -" },
            { "RPL_ENDOFMOTD", "{0} :End of /MOTD" },
            { "RPL_OPER", "{0} :Operator privileges granted" },
            { "RPL_SERVICE", "{0} {1} :Service privileges granted" },
            { "RPL_TIME", "{0} {1} {2} {3} :{4}" },
            { "ERR_UNKNOWN", "{0} {1} :{2}" },
            { "ERR_NOSUCHNICK", "{0} {1} :No such nick" },
            { "ERR_NOSUCHSERVER", "{0} {1} :No such server" },
            { "ERR_NOSUCHCHANNEL", "{0} {1} :No such channel" },
            { "ERR_CANNOTSEND", "{0} {1} :Cannot send to channel" },
            { "ERR_TOOMANYCHANNELS", "{0} {1} :You have joined too many channels" },
            { "ERR_TOOMANYTARGETS", "{0} :{1}" },
            { "ERR_NOSUCHSERVICE", "{0} {1} :No such service" },
            { "ERR_NOORIGIN", "{0} :No origin specified" },
            { "ERR_INVALIDCAP", "{0} {1} :Invalid CAP command" },
            { "ERR_TOOMANYMATCHES", "{0} {1} :There are too many matches; please refine your query" },
            { "ERR_UNKNOWNCOMMAND", "{0} {1} :Unknown command" },
            { "ERR_NOMOTD", "{0} :Message of the Day is not available" },
            { "ERR_NONICKNAMEGIVEN", "{0} :{1}" },
            { "ERR_BADNICK", "{0} {1} :Nickname is invalid ({2})" },
            { "ERR_NICKNAMEINUSE", "{0} {1} :Nickname is already in use" },
            { "ERR_UNAVAILRESOURCE", "{0} {1} :Nick/channel is temporary unavailable" },
            { "ERR_USERNOTINCHANNEL", "{0} {1} {2} :Target not in channel" },
            { "ERR_NOTONCHANNEL", "{0} {1} :Not on channel" },
            { "ERR_NOTREGISTERED", "{0} :You are not registered" },
            { "ERR_NEEDPARAMS", "{0} {1} :Not enough parameters" },
            { "ERR_REGISTERED", "{0} :You are already authorized" },
            { "ERR_PASSWDMISMATCH", "{0} :The password you provided does not match" },
            { "ERR_BANNED", "{0} :You are banned" },
            { "ERR_INVALIDUSERNAME", "{0} {1} :Invalid username" },
            { "ERR_CHANNELISFULL", "{0} {1} :Channel is full (+l)" },
            { "ERR_UNKNOWNMODE", "{0} {1} :is not a recognized mode" },
            { "ERR_INVITEONLYCHAN", "{0} {1} :Channel is invite-only (+i)" },
            { "ERR_BANNEDFROMCHAN", "{0} {1} :You may not join this channel" },
            { "ERR_BADCHANNELKEY", "{0} {1} :You may not join this channel" },
            { "ERR_BADCHANMASK", "{0} {1} :Bad channel mask" },
            { "ERR_NOCHANMODES", "{0} {1} :Channel does not support modes" },
            { "ERR_BANLISTFULL", "{0} {1} {2} :Channel list is full" },
            { "ERR_BADCHANNAME", "{0} {1} :Invalid channel name ({2})" },
            { "ERR_NEEDACCOUNT", "{0} {1} :Channel is restricted to authenticated users" },
            { "ERR_NOTOPER", "{0} :Only members of Operations staff can do that" },
            { "ERR_CHANOPRIVSNEEDED", "{0} {1} :You are not a channel operator" },
            { "ERR_CANTKILLSERVER", "{0} :You cannot /KILL a server" },
            { "ERR_ISSERVICE", "{0} {1} :Services may not be kicked, deopped, or killed" },
            { "ERR_NOOPERHOST", "{0} :Your host is not authorized" },
            { "ERR_UMODEUNKNOWN", "{0} :Unknown MODE flag {1}" },
            { "ERR_USERSDONTMATCH", "{0} :MODE may only be set for yourself" },
            { "RPL_OMOTDSTART", "{0} :- Oper Message of the Day -" },
            { "RPL_OMOTD", "{0} :- {1}" },
            { "RPL_ENDOFOMOTD", "{0} :End of Oper MOTD" },
            { "ERR_NOPRIVS", "{0} :Authorization failed" },
            { "RPL_MONONLINE", "{0} :{1}" },
            { "RPL_MONOFFLINE", "{0} :{1}" },
            { "RPL_MONLIST", "{0} :{1}" },
            { "RPL_ENDOFMONLIST", "{0} :End of MONITOR list" },
            { "ERR_MONLISTFULL", "{0} {1} {2} :Monitor list is full" },
            { "ERR_INVALIDBAN", "{0} {1} {2} {3} :Invalid ban/except mask" },
            { "RPL_LOGGEDIN", "{0} {1} {2} :You are now logged in as {2}" },
            { "RPL_LOGGEDOUT", "{0} {1} :You are now logged out" },
        };

        static void Main(string[] args)
        {
            var doc = XmlSpreadsheetFile.Workbook.LoadFromFile(@"E:\XmlSpreadsheetParser\test\inventory.xml");
            var sheet = doc.Worksheets[0];
            Console.Write($"Worksheets.Count: {doc.Worksheets.Count}");

            doc.SaveToFile(@"E:\XmlSpreadsheetParser\test\inventory2.xml");

            var test = sheet.Table.AsDataTable();

            var data = new System.Data.DataTable();

            data.Columns.Add("Number", typeof(int));
            data.Columns.Add("Identifier");
            data.Columns.Add("Text");

            int n = 1;
            foreach (var item in testValues) {
                data.Rows.Add(n++, item.Key, item.Value);
            }

            sheet.Table = XmlSpreadsheetFile.Table.FromDataTable(data);

            doc.DocumentProperties.LastSaved = DateTime.UtcNow;
            doc.WorkbookProperties.WindowHeight = null;
            doc.WorkbookProperties.WindowWidth = null;
            doc.WorkbookProperties.WindowTopX = null;
            doc.WorkbookProperties.WindowTopY = null;
            doc.WorkbookProperties.ProtectStructure = true;
            doc.SaveToFile(@"E:\XmlSpreadsheetParser\test\test2.xml");
        }
    }
}
