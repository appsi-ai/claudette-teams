Teams app package
=================

This folder is zipped into manifest.zip and sideloaded into Teams.

Required contents of manifest.zip:
  - manifest.json
  - color.png    (192 x 192)
  - outline.png  (32 x 32, transparent)

Before zipping, in manifest.json replace:
  <<AAD_APP_CLIENT_ID>>  -> your Azure Bot resource App (client) ID
  <<BOT_DOMAIN>>         -> your App Service host, e.g. claudette-teams.azurewebsites.net

Add the two PNG icons here (placeholders are not committed). Then:
  PowerShell: Compress-Archive -Path manifest.json,color.png,outline.png -DestinationPath manifest.zip

Upload via Teams Admin Center -> Teams apps -> Manage apps -> Upload new app,
or for testing: Teams -> Apps -> Manage your apps -> Upload a custom app.
