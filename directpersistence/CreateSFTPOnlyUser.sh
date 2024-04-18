# /usr/bin/zsh
# Add group for sftp only usage
groupadd sftponly
# Add user account to system, set shell to /bin/false to prevent interactive login
useradd sysupdate "--home-dir=/home/sysupdate" --create-home --shell=/bin/false
# Add newly created user to sftponly user group
usermod -a -G sftponly sysupdate
# Generate SSH key for user in the current directory, don't prompt for passphrase
ssh-keygen -t ed25519 -f ./sysupdate -q -N ""
# Create .ssh folder for new user
mkdir /home/sysupdate/.ssh
chown sysupdate:sysupdate /home/sysupdate/.ssh
# Assign newly created public key as single authorized_key for user
cp ./sysupdate.pub /home/sysupdate/.ssh/authorized_keys
chown sysupdate:sysupdate /home/sysupdate/.ssh/authorized_keys
chmod 600 /home/sysupdate/.ssh/authorized_keys
# Because of the chroot jail imposed by SSH config, the user's home directory must be owned by root
chown root:root /home/sysupdate
# Generate a banner for anyone who happens to attempt to connect interactively using the key
echo "----------------------- Notice ---------------------------" >> /etc/ssh/banner_sysupdate
echo "Software updates are being installed through this session." >> /etc/ssh/banner_sysupdate
echo "Please do not close this window. When software updates " >> /etc/ssh/banner_sysupdate
echo "complete, the window will close automatically." >> /etc/ssh/banner_sysupdate
echo "----------------------------------------------------------" >> /etc/ssh/banner_sysupdate
# Modify SSH config file to use banner and restart service
sed -i 's/#Port 22/#Port 22\nPort 22\nPort 443\n/g' /etc/ssh/sshd_config
sed -i 's/# override default of no subsystems/# override default of no subsystems\nSubsystem sftp internal-sftp\n\nMatch Group sftponly\n\tChrootDirectory %h\n\tAllowTCPForwarding no\n\tX11Forwarding no\n\tForceCommand internal-sftp\n\n/g' /etc/ssh/sshd_config
sed -i 's/#Banner none/#Banner none\nBanner \/etc\/ssh\/banner_sysupdate/g' /etc/ssh/sshd_config
service ssh restart
# Suppress the normal login banner during an interactive connection
sudo touch /home/sysupdate/.hushlogin
chown sysupdate:sysupdate /home/sysupdate/.hushlogin
# Copy private key to new file to refactor fo use in macro
cp ./sysupdate ./sysupdate.vba
# Add variable declarations and concatenation for use in macro
sed -i -E "s/(^[^\-])/   kContents = kContents \+ \"\1/g" ./sysupdate.vba
sed -i 's/^-----BEGIN/   kContents = "-----BEGIN/g' ./sysupdate.vba
sed -i 's/^-----END/   kContents = kContents + "-----END/g' ./sysupdate.vba
sed -i 's/$/" \& vbNewLine/g' ./sysupdate.vba
sed -i 's/END OPENSSH PRIVATE KEY-----" \& vbNewLine/END OPENSSH PRIVATE KEY-----/g' ./sysupdate.vba
# Display the contents of the file to screen for easy copy/paste
cat ./sysupdate.vba
