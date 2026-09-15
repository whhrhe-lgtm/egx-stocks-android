[app]

title = EGX Stocks
package.name = egxstocks
package.domain = org.egxstocks

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx

version = 1.0

requirements = python3,kivy,openpyxl

orientation = portrait
fullscreen = 0

services =

presplash.filename =
icon.filename =

android.archs = arm64-v8a, armeabi-v7a
android.api = 36
android.minapi = 24
android.ndk = 28c

p4a.fork = kivy
p4a.branch = develop
p4a.commit = HEAD
p4a.bootstrap = sdl2

android.permissions = INTERNET

android.add_src =

android.entrypoint = org.kivy.android.PythonActivity
android.activity_class_name = org.kivy.android.PythonActivity
android.apptheme = "@android:style/Theme.Material.Light.NoActionBar"

log_level = 2
presplash_color = #FFFFFF

[buildozer]

log_level = 2
warn_on_root = 1

android.gradle_dependencies =
