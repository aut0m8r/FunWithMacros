## Direct Persistence Resources

This project consists of files that can be used to configure an SSH user on a VPS for SFTP or tunnel-only access. Corresponding VBA macro examples are included to generate ideas regarding unique Macro payloads.

### SFTP Payload Retrieval

In the SFTP case, the CreateSFTPOnlyUser.sh script is used on the server. The output from the script is placed in the Shortcut_Macro_DLL_Hijack.txt macro source, replacing the SSH private key. The outbound SFTP connection is used to retrieve a payload for execution. The corresponding Office macro retrieves the payload via SFTP, then drops it into a common DLL hijack location. This can be easily modified to retrieve a payload, drop an LNK file to a LOLBin that references the file, and results in malware execution when the user logs onto the host. The file retrieved using SFTP does not get mark of the web.

#### SSH Tunnel Access

In the SSH tunnel case, the CreateTunnelOnlyUser.sh script is used on the server. The output from the script is placed in the Shortcut_Macro_SSH_Tunnel.txt macro source, replacing the SSH private key. The outbound SSH connection is used to create a reverse tunnel into the target network, allowing stable connectivity and easy transport of external tooling traffic.
