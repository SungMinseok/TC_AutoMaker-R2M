import os
import paramiko

# Set variables
ftp_host = "nas.webzen.co.kr"
ftp_port = 23145
ftp_user = "mssung@webzen.com"
ftp_password = "webzen@2301"
local_folder = "D:\UploadTest"
ftp_folder = "/"

# Connect to FTP server
transport = paramiko.Transport((ftp_host, ftp_port))
transport.connect(username=ftp_user, password=ftp_password)
sftp = paramiko.SFTPClient.from_transport(transport)

# Change to FTP directory
sftp.chdir(ftp_folder)

# Upload files
for filename in os.listdir(local_folder):
    local_path = os.path.join(local_folder, filename)
    if os.path.isfile(local_path):
        sftp.put(local_path, filename)

# Close connection
sftp.close()
transport.close()