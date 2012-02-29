# Envelope

Some time after the release of Adobe ColdFusion 9, Microsoft released Exchange 2010. Until this point, ColdFusion had functionality for interacting with Exchange, but it was only compatible up to Exchange 2007.

This project is a wrapper that we wrote to access some of the Exchange functionality via EWS (Exchange Web Services). Unfortunately, EWS is so complex and convoluted that attempting to interact with it directly from anything other than .NET proved fruitless; so we decided to create a .NET wrapper API, and then a ColdFusion CFC that would work with that API. It's a little fragile, for sure, but it works. For us, it's done a decent enough job of bridging the gap between the release of Exchange 2010 and ColdFusion 10 (which is now compatible via CFExchange functionality again), so we thought we would share our solution.

Hopefully it will help you in some meaningful way.

## Impersonation

Envelope works by connecting to EWS as an account with permissions to "impersonate" any exchange user, and then impersonating them to complete the requested action. The way we've used it is to create 1 impersonation account per application that uses this API, for auditing purposes. You'll then need to configure Web.config (in the .NET app) to have an API key for each application to use, and set the username and password associated with each API key. This is not the most elegant solution, but on a small scale it works well enough.

## SQL

We log all activity to an MSSQL database. Scripts to create the log table and stored procedure are included in the SQL folder.

## Configuration

You'll need to update the code in the following locations to be specific to your environment:

1. ./dotnet/App_Code/WhartonEWS.cs line 49 -- update for your exchange environment
1. ./dotnet/web.config -- lines 27-55 -- generate API keys that you want to use and add appropriate items here. See "API Keys" section below for more information.
1. ./dotnet/web.config -- line 72 -- configure connection string for logging (required)
1. ./CFC/Envelope_v1.cfc -- line 6 -- update for the location where you placed Envelope's .NET code

### API Keys

To create API keys, we simply went to http://createguid.com/ and copied the created GUID, removing dashes.

There are 4 pieces to each API key:

1. List of all API keys, for easy reference (the code doesn't use these at all). Lines 32-35.
1. Impsersonator account username
1. Impersonator account password
1. Domain that the impersonator account belongs to

There's nothing stopping you from using the same impersonator account for every api key, but we separated them for extra audit abilities, since Exchange logs which impersonator performs any actions, in addition to our own baked in logging.

# LICENSE

This code is licensed under the MIT LICENSE. See the LICENSE file for more information.
