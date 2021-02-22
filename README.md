# Exchange Online PowerShell interface

This module makes it easier for Node.js applications to communicate with Exchange Online via PowerShell.

Tested in Kubuntu 20.10, Node.js 14.

## Setup

The server must have PowerShell and Exchange Online PowerShell V2 module version 2.0.4 or newer installed.

The patched [OMI](https://github.com/jborean93/omi) must be installed as well.

You'd need to create an app in Azure AD with a PFX certificate generated and uploaded.

```shell
npm i @keleran/exon
```

## Usage

```javascript
const Exon = require('@keleran/exon');

const exon = new Exon({
  // mandatory
  connection: {
    appId: '*',
    // path to the .pfx file
    certificateFilePath: '*',
    // account domain 
    organization: '*.onmicrosoft.com'
  },
  // optional, defaults to the shown values
  timeouts: {
    spawn: 5000,
    exec: 180000
  },
  // optional - where to write log entries to 
  logFilePath: '*'
});

// resolved with stdout content
// rejected with stderr content
// there's no queue, wait for it to finish before sending a new command or it will be rejected instantly
// the first run will take extra time to download cmdlets from the remote server
const result = await exon.exec('CMDLET_NAME', {'CMDLET_PARAM_NAME': 'CMDLET_PARAM_VALUE'});

// disconnect from Exchange and kill PS process
// it will respawn and reconnect automatically on the next .exec()
// useful for graceful shutdown or if you believe it could help with some PS-related issue 
await exon.reset();
```