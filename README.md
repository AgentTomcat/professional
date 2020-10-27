# professional

Welcome, all! Here begins the collection of code snippets I've created over the years to help myself or others.

I've used Microsoft VBA to solve myriad problems, so most code snippets will probably be in that language, but some may be in AHK or Windows scripts.

One notable project I worked on was a prodictivity tracking and timekeeping system that was written in AHK. The back-end reporting section of it relied on Access, Excel, and Windows scripts to pull everything together and produce useful reports.
I can't share it here for legal reasons, but here's a brief rundown:
This system had to track time punches accurately, with error handling and built-in redundancy to make up for system errors and human errors or subterfuge.
It had to do so with an easy interface that was easy to understand with little or no instruction.
It had to provide a visible break timer to help users know track their breaks and lunches.
It also, quite separately, had to provide a means by which a user could input completion data for individual tasks, along with associated time tracking data (started task, ended task, paused task, etc.).
It had to maintain a local and remote crash tracking and backup task tracking log for instances when automatic error compensation couldn't work.
It had to do all of this quickly, without frustrating the user or causing unnecessary delay.

In the end, it achieved all of this. Any delays were from the system itself, usually the site server, where the program tracked that it sent the data, but the server never received it. This was fortunately rare, because fixes for data loss were manual.
