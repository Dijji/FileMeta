Two major users of the property handler when looking at a folder are Explorer and the Indexing Server. It is important to minimise collisions between these two in order that both behave as well as possible.

This test program simulates the behaviours of Explorer and the Indexing Server, pounding on the files in a single test folder trying to generate collisions.

As a stand-alone program, it can take advantage of some of the evolution of C# since File Meta was first created. It is best opened using Visual Studio 2019 or later.

The program runs two threads. The Explorer thread reads each file in the folder, reads the Comments property from it, and updates the Comments every 20th file. The Indexer reads every property from each file in the folder.

Both iterate until the slower thread has read the requested number of files (the default is 10,000). Both threads run flat out: a run with 10,000 files might take around a minute. A log file is produced, with a variable level of detail, that might look something like this:

Started at 31/05/2020 11:57:56
Indexer: 13048 files, 47 failures, 0 terminated. Explorer: 10001 files, 71 failures, 0 terminated. 
Run took 00:00:53.7582314
Collision rate was 0.59%

That was actually taken from a better than average run for release 1.6.0.4. Taken over a series of runs, representative figures for the collision rates for different releases are:

1.6      1.3%
1.6.0.3  2.6%
1.6.0.4  1.1%

As can be seen, 1.6.0.3 behaves considerably worse than the current recommended release, but 1.6.0.4 recovers the lost ground, plus a little more.

However, remember that this program deliberately pushes stress to the extremes: none of these figures should have caused any problems to manifest in day-to-day usage of File Meta.