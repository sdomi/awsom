# awsom

Awsom is a GNU/Social / Mastodon / Pleroma client written from scratch in (sic!) VB6. It's target are retro computers which can handle displaying a feed of messages with or without images, but can't properly render the web UI. Currently, awsom will run on everything with OS greater than Windows 2000, and *probably* on Windows 98SE with KernelEx, but I'm working on a version that will support everything from Windows 95 up.

To run the code, you'll need curl binary - obtain yours from [curl's website](https://curl.haxx.se/), or get it from the latest awsom's binary release. In the future, I plan to add wget and/or libcurl support as well.

Awsom is in active development state - feel free to contribute.

### Working features
* Displaying a feed
* Sending toots
* GUI form for setting up the client (API token, instance name)

### Broken/WIP features
* Likes (code needs a rewrite, because event handling sucks)
* Retoots
* Replies
* Notifications
* Instance feed
* Federated feed
* Images/Videos
* URLs
* parsing &apos, &lt, &gt and a few others
* icons (they'll come.)
* other cool stuff