PowerPad
========
PowerPad is a simple and easy way to replicate the functionality of Powerpoints presenters view, though on your tablet or phone.

Features
--------
* Supports Powerpoint 2013.
* Interface is available through any device with a browser - including iPads, iPhones as well as Android & Windows tablets & Phones.
* Automatically caches slides & detects changes when presentation starts.

PowerPad Interface
------------------
Instead of replicating the Powerpoint presenters view 1:1, I've aimed for showing just what's needed, and no more.

* At the back there's a full-size preview of the next slide.
* At the bottom you'll see the notes for the current slide (with [Markdown](http://daringfireball.net/projects/markdown/syntax) support).
* At the top there are two progress bars, the top one showing your progress in time, the bottom one showing your progress in slides. Also, you'll see the current time as well as the presentation begin & end times here (both configurable - just press them).

![Notes](/readme/screens/screen_notes.png)

Optimally you'll want to keep the two progress bars in sync (unless you have few slides that take disproportionally long time). If this is your situation, you're running behind.

![Running behind](/readme/screens/screen_behind.png)

Once you reach the end, PowerPad will notify you.

![The end](/readme/screens/screen_end_of_slideshow.png)

If you don't have a tablet by hand, you may even use the interface on your cell phone, though a larger display is preferable.

![The end](/readme/screens/screen_mobile.png)

Getting Started
---------------
As soon as you start up PowerPad, it'll detect whether you're already running Powerpoint. If an instance is detected, PowerPad will conect to it and await for the presentation to begin.

![Starting PowerPad with existing Powerpoint instance running](/readme/screens/just_started_existing_powerpoint.png)

If Powerpoint is not already running, Powerpad will launch it for you.

![Starting PowerPad](/readme/screens/just_started_no_powerpoint.png)

As can be seen in the green output, PowerPad will listen on any active IP address on port 8000, by default. At this point, as soon as you begin the presentation in Powerpoint, PowerPad will run through all slides and cache them.

![Starting presentation](/readme/screens/presentation_started.png)

Caching is a relatively quick process. It will however lock up Powerpoint while it's running, but as soon as it's done, you've got full control over Powerpoint and PowerPad will automatically detect which slide is the active one.

![Changing slides](/readme/screens/presentation_changing_slides.png)

If you end the slideshow, perhaps for a demo, and restart it afterwards, PowerPad will detect this and continue its work in the background. If any slides have been changed in the meantime, they'll be cached again. Unchanged slides will be ignored, making this a very quick process.

![Restarting presentation](/readme/screens/presentation_restarted.png)

Frequently Asked Questions
--------------------------
### How do I configure what port PowerPad will listen on?
Simply open up PowerPad.exe.config and change this line:

    <appSettings>
		<add key="Port" value="8000"/>
    </appSettings>
    
Make sure to open up for the port in your firewall!

### How do I set the presentation begin & end times?
Simply tap on the times and enter them in 24H format like ##:## - 05:30, 14:15, 17:00, etc. PowerPad will set a cookie and remember the values even if you refresh the window.

### Can I use multiple clients at the same time?
Yes! PowerPad will handle any number of clients so you can hook up 15 phones and 5 tablets, should you want to. You can even give your attendees access!

### How do I format my notes?
PowerPad does not support the native rich text formatting that PowerPoint stores, but it will render Markdown. If you just write your notes in Markdown format, they will be shown as such in the PowerPad interface.

### Why perform caching up front and not on-demand?
Making Powerpoint render a slide locks up Powerpoint for just a short moment. Unfortunately this is enough to cause some problems with clickers and changing slides using the keyboard. As such, aggressively caching seems to be the only stable method.

Requirements
------------
* Powerpoint 2013
* .NET Framework 2.0
* Firewall setup to allow clients on the configured port (8000 by default).

Future
------
One of the design guidelines for PowerPad is to keep it as absolutely simple as possible, though no simpler than that. Given that, I still have a couple of features in mind that I'd like to add (feel free to help, if you feel like it).

* Support for at least Powerpoint 2010. This should be simple as the interop functionality is pretty much the same.
* Multiple views that can be access by just invoking a special URL.
 * Display of current slide and nothing else.
 * Display of next slide and nothing else.
 * More or less exact replica of Powerpoints presenters view - good if you have a secondary laptop for showing the interface.

Contact
-------
For any questions, issues or suggestions, please contact me at

* Mail: mark@improve.dk
* Twitter: [@improvedk](https://twitter.com/improvedk)
* Blog: [improve.dk](http://improve.dk/)