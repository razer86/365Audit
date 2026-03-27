# Invoke-TeamsAudit.ps1

Connects to Microsoft Teams via the MicrosoftTeams module and audits federation, client configuration, meeting policies, guest access, and app policies.

## Required Permissions

Granted automatically by `Setup-365AuditApp.ps1`:

| Permission | Type | Purpose |
|---|---|---|
| `TeamsAppInstallation.ReadForUser.All` | Application | App permission and setup policies |
| `TeamMember.Read.All` | Application | Team membership data |

> The Teams module also uses the existing Graph and Exchange connections established earlier in the audit session.

## Output Files

| File | Description |
|---|---|
| `Teams_FederationConfig.csv` | External access (federation) settings — allowed/blocked domains, Teams and Skype federation flags |
| `Teams_ClientConfig.csv` | Teams client configuration — file sharing, cloud storage, external app, and communication settings |
| `Teams_MeetingPolicies.csv` | Meeting policies — recording, transcription, lobby, and external participant settings |
| `Teams_GuestMeetingConfig.csv` | Guest meeting configuration — IP audio/video, screen share, and meeting flags |
| `Teams_GuestCallingConfig.csv` | Guest calling settings |
| `Teams_MessagingPolicies.csv` | Messaging policies — Giphy, memes, URL previews, read receipts, and chat edit/delete settings |
| `Teams_AppPermissionPolicies.csv` | App permission policies — Microsoft, third-party, and tenant app access controls |
| `Teams_AppSetupPolicies.csv` | App setup policies — pinned apps and user pinning settings |
| `Teams_ChannelPolicies.csv` | Channel policies — private and shared channel creation permissions (skipped gracefully if not available in module version) |
