#!/bin/bash

gsettings set org.gnome.desktop.interface scaling-factor 2
gsettings set org.gnome.settings-daemon.plugins.xsettings overrides "{'Gdk/WindowScalingFactor': <2>}"
xrandr --output HDMI-0 --scale 1.50x1.50
#xrandr --output HDMI-0 --panning 3840x2160
