# !/bin/bash
# Mount G: drive to /mnt/g if not already mounted
# This is useful for WSL users who want to access their Windows drives from the Linux environment. The script checks if /mnt/g is already a mount point, and if not, it uses the `mount` command to mount the G: drive to /mnt/g using the drvfs filesystem type. This allows you to access your G: drive files from within WSL.
if ! mountpoint -q /mnt/g; then
  sudo mount -t drvfs G: /mnt/g
fi