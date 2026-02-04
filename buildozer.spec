[app]

# (str) Title of your application
title = Student Tracker Pro

# (str) Package name
package.name = studenttrackerpro

# (str) Package domain (needed for android/ios packaging)
package.domain = com.younesbennacer

# (str) Source code where the main.py live
source.dir = .

# (list) Source files to include (let empty to include all the files)
source.include_exts = py,png,jpg,kv,atlas,db,txt

# (str) Application versioning (method 1)
version = 2.0

# (list) Application requirements
# comma separated e.g. requirements = sqlite3,kivy
requirements = python3,kivy==2.3.0,pandas,numpy,openpyxl,android,jnius,pyjnius

# (str) Supported orientation (landscape, portrait or all)
orientation = landscape

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

# (list) Permissions
# CRITICAL: These permissions enable the settings UI in Android
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# (int) Target Android API, should be as high as possible.
# For Android 11+ (API 30+), we need this for scoped storage
android.api = 33

# (int) Minimum API your APK will support.
android.minapi = 21

# (int) Android SDK version to use
android.sdk = 33

# (str) Android NDK version to use
android.ndk = 25b

# (bool) Use --private data storage (True) or --dir (False)
android.private_storage = True

# (str) Android NDK directory (if empty, it will be automatically downloaded.)
#android.ndk_path =

# (str) Android SDK directory (if empty, it will be automatically downloaded.)
#android.sdk_path =

# (str) ANT directory (if empty, it will be automatically downloaded.)
#android.ant_path =

# (bool) If True, then skip trying to update the Android sdk
# This can be useful to avoid excess Internet downloads or save time
# when an update is due and you just want to test/build your package
android.skip_update = False

# (bool) If True, then automatically accept SDK license
# agreements. This is intended for automation only. If set to False,
# the default, you will be shown the license when first running
# buildozer.
android.accept_sdk_license = True

# (str) Android entry point, default is ok for Kivy-based app
#android.entrypoint = org.kivy.android.PythonActivity

# (str) Full name including package path of the Java class that implements Android Activity
# use that parameter together with android.entrypoint to set custom Java class instead of PythonActivity
#android.activity_class_name = org.kivy.android.PythonActivity

# (str) Extra xml to write directly inside the <manifest> element of AndroidManifest.xml
# use that parameter to provide a filename from where to load your custom XML code
#android.extra_manifest_xml = ./src/android/extra_manifest.xml

# (str) Extra xml to write directly inside the <application> element of AndroidManifest.xml
# use that parameter to provide a filename from where to load your custom XML code
#android.extra_manifest_application_xml = ./src/android/extra_manifest_application.xml

# (list) Pattern to whitelist for the whole project
#android.whitelist =

# (str) Path to a custom whitelist file
#android.whitelist_src =

# (str) Path to a custom blacklist file
#android.blacklist_src =

# (list) List of Java .jar files to add to the libs so that pyjnius can access
# their classes. Don't add jars that you do not need, since extra jars can slow
# down the build process. Allows wildcards matching, for example:
# OUYA-ODK/libs/*.jar
#android.add_jars = foo.jar,bar.jar,path/to/more/*.jar

# (list) List of Java files to add to the android project (can be java or a
# directory containing the files)
#android.add_src =

# (list) Android AAR archives to add
#android.add_aars =

# (list) Put these files or directories in the apk assets directory.
# Either form may be used, and assets need not be in 'source.include_exts'.
# 1) android.add_assets = source_dir/image.png
# 2) android.add_assets = source_dir/images/
#android.add_assets =

# (list) Put these files or directories in the apk res directory.
# The option may be used in three ways, the value may contain one or zero ':'
# Some examples:
# 1) A file to add to resources, legal resource names contain ['a-z','0-9','_']
# android.add_resources = legal_name:path/to/file1,legal_name2:path/to/file2
# 2) A directory, here  'legal_resource_name' should be a valid file name in the directory
# android.add_resources = path/to/directory1,path/to/directory2
# 3) A directory with custom resource names:
# android.add_resources = path/to/directory1:drawable-hdpi/,path/to/directory2:drawable-xhdpi/
#android.add_resources =

# (list) Gradle dependencies to add
#android.gradle_dependencies =

# (bool) Enable AndroidX support
android.enable_androidx = True

# (list) add java compile options
# this can for example be necessary when importing certain java libraries using the 'android.gradle_dependencies' option
# see https://developer.android.com/studio/write/java8-support for further information
# android.add_compile_options = "sourceCompatibility = 1.8", "targetCompatibility = 1.8"

# (list) Gradle repositories to add {can be necessary for some android.gradle_dependencies}
# please enclose in double quotes 
# e.g. android.gradle_repositories = "maven { url 'https://kotlin.bintray.com/ktor' }"
#android.add_gradle_repositories =

# (str) The Android arch to build for, choices: armeabi-v7a, arm64-v8a, x86, x86_64
# You can build for multiple archs at the same time, just separate with a comma
android.archs = arm64-v8a, armeabi-v7a

# (bool) enables Android auto backup feature (Android API >=23)
android.allow_backup = True

# (str) The format used to package the app for release mode (aab or apk).
# android.release_artifact = aab

# (str) The format used to package the app for debug mode (apk or aab).
# android.debug_artifact = apk

#
# Python for android (p4a) specific
#

# (str) python-for-android URL to use for checkout
#p4a.url =

# (str) python-for-android fork to use in case if p4a.url is not specified, defaults to upstream (kivy)
#p4a.fork = kivy

# (str) python-for-android branch to use, defaults to master
#p4a.branch = master

# (str) python-for-android specific commit to use, defaults to HEAD, must be within p4a.branch
#p4a.commit = HEAD

# (str) python-for-android git clone directory (if empty, it will be automatically cloned from github)
#p4a.source_dir =

# (str) The directory in which python-for-android should look for your own build recipes (if any)
#p4a.local_recipes =

# (str) Filename to the hook for p4a
#p4a.hook =

# (str) Bootstrap to use for android builds
# p4a.bootstrap = sdl2

# (int) port number to specify an explicit --port= p4a argument (eg for bootstrap flask)
#p4a.port =

# Control passing the --use-setup-py vs --ignore-setup-py to p4a
# "in the future" --use-setup-py is going to be the default behaviour in p4a, right now it is not
# Setting this to false will pass --ignore-setup-py, true will pass --use-setup-py
# NOTE: this is general setuptools integration, having pyproject.toml is enough, no need to generate
# setup.py if you're using Poetry, but you need to add "toml" to source.include_exts.
#p4a.setup_py = false

# (str) Launcher icon (in the res/drawable directory)
#icon.filename = %(source.dir)s/data/icon.png

# (str) Launcher icon for adaptive Android app
# This icon will have a transparent background and will be foreground icon
#icon.adaptive_foreground.filename = %(source.dir)s/data/icon_fg.png

# (str) Launcher icon for adaptive Android app
# This icon will be the background layer and must be a solid color
#icon.adaptive_background.filename = %(source.dir)s/data/icon_bg.png

# (list) Presplash image background color (for Android toolchain)
# Supported formats are: #RRGGBB #AARRGGBB or one of the following names:
# red, blue, green, black, white, gray, cyan, magenta, yellow, lightgray,
# darkgray, grey, lightgrey, darkgrey, aqua, fuchsia, lime, maroon, navy,
# olive, purple, silver, teal.
#android.presplash_color = #FFFFFF

# (str) Presplash animation using Lottie format.
# see https://lottiefiles.com/ for examples and https://airbnb.design/lottie/ for general documentation.
# Lottie files can be created using various tools, like Adobe After Effects or Synfig.
#android.presplash_lottie = "path/to/lottie/file.json"

# (str) Adaptive icon background color (for Android toolchain)
#icon.adaptive_background.color = #FFFFFF

# (list) Permissions app will request
# Standard permissions (automatically granted):
# - INTERNET
# - ACCESS_NETWORK_STATE
# - WAKE_LOCK
# Runtime permissions (require user approval):
# - READ_EXTERNAL_STORAGE
# - WRITE_EXTERNAL_STORAGE
# Special permissions (require manual enabling in settings for Android 11+):
# - MANAGE_EXTERNAL_STORAGE

# IMPORTANT ANDROID 11+ (API 30+) CONFIGURATION
# For scoped storage and proper file access on Android 11+, we need:
# 1. MANAGE_EXTERNAL_STORAGE permission (declared above)
# 2. requestLegacyExternalStorage flag in manifest (for Android 10 compatibility)
# 3. Proper runtime permission requests in code (already in main.py)

# (str) XML to be included in AndroidManifest.xml
android.manifest.intent_filters = 

# (str) launchMode to set for the main activity
#android.manifest.launch_mode = standard

# (list) Android additionnal libraries to copy into libs/armeabi
#android.add_libs_armeabi = libs/android/*.so
#android.add_libs_armeabi_v7a = libs/android-v7/*.so
#android.add_libs_arm64_v8a = libs/android-v8/*.so
#android.add_libs_x86 = libs/android-x86/*.so
#android.add_libs_mips = libs/android-mips/*.so

# (bool) Copy library instead of making a libpymodules.so
#android.copy_libs = 1

# (list) The Android application meta-data to set (key=value format)
#android.meta_data =

# (str) Android logcat filters to use
#android.logcat_filters = *:S python:D

# (bool) Android logcat only display log for activity's pid
#android.logcat_pid_only = False

# (str) Android additional adb arguments
#android.adb_args = -H host.docker.internal

# (bool) Copy library instead of making a libpymodules.so
#android.copy_libs = 1

# (list) The Android application meta-data to set (key=value format)
# For Android 10 (API 29) compatibility with legacy external storage
android.meta_data = android.app.requestLegacyExternalStorage=true

# (str) Android backup rules file (for Android Auto Backup)
#android.backup_rules =

#
# iOS specific
#

# (str) Path to a custom kivy-ios folder
#ios.kivy_ios_dir = ../kivy-ios
# Alternately, specify the URL and branch of a git checkout:
ios.kivy_ios_url = https://github.com/kivy/kivy-ios
ios.kivy_ios_branch = master

# Another platform dependency: ios-deploy
# Uncomment to use a custom checkout
#ios.ios_deploy_dir = ../ios_deploy
# Or specify URL and branch
ios.ios_deploy_url = https://github.com/phonegap/ios-deploy
ios.ios_deploy_branch = 1.10.0

# (bool) Whether or not to sign the code
ios.codesign.allowed = false

# (str) Name of the certificate to use for signing the debug version
# Get a list of available identities: buildozer ios list_identities
#ios.codesign.debug = "iPhone Developer: <lastname> <firstname> (<hexstring>)"

# (str) The development team to use for signing the debug version
#ios.codesign.development_team.debug = <hexstring>

# (str) Name of the certificate to use for signing the release version
#ios.codesign.release = %(ios.codesign.debug)s

# (str) The development team to use for signing the release version
#ios.codesign.development_team.release = <hexstring>

# (str) URL pointing to .ipa file to be installed
# This option should be defined along with `display-image-url` and `full-size-image-url` options.
#ios.manifest.app_url =

# (str) URL pointing to an icon (57x57px) to be displayed during download
# This option should be defined along with `app-url` and `full-size-image-url` options.
#ios.manifest.display_image_url =

# (str) URL pointing to a large icon (512x512px) to be used by iTunes
# This option should be defined along with `app-url` and `display-image-url` options.
#ios.manifest.full_size_image_url =

#
# OSX Specific
#

# (str) The name of the application bundle
#osx.bundle_name = StudentTrackerPro

# (str) The id of the bundle
#osx.bundle_id = com.younesbennacer.studenttrackerpro

# (str) The version of the bundle
#osx.bundle_version = 2.0

# (str) The build number of the bundle
#osx.bundle_build = 1

#
# Buildozer
#

# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root (0 = False, 1 = True)
warn_on_root = 1

# (str) Path to build artifact storage, absolute or relative to spec file
# build_dir = ./.buildozer

# (str) Path to build output (i.e. .apk, .aab, .ipa) storage
# bin_dir = ./bin

#-----------------------------------------------------------------------------
# List as sections
#
# You can define all the "list" as [section:key].
# Each line will be considered as a option to the list.
# Let's take [app] / source.exclude_patterns.
# Instead of doing:
#
#     [app]
#     source.exclude_patterns = license,data/audio/*.wav,data/images/original/*
#
# This can be translated into:
#
#     [app:source.exclude_patterns]
#     license
#     data/audio/*.wav
#     data/images/original/*
#

#-----------------------------------------------------------------------------
# Profiles
#
# You can extend section / key with a profile
# For instance, you want to deploy a demo version of your application without
# HD content. You could first change the title to add "(demo)" in the name
# and extend the excluded directories to remove the HD content.
#
#     [app@demo]
#     title = My Application (demo)
#
#     [app:source.exclude_patterns@demo]
#     images/hd/*
#
# Then, invoke the command line with the "demo" profile:
#
#     buildozer --profile demo android debug
