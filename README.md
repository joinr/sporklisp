sporklisp
=========

A horrifying experiment in implementing a lisp in VBA.  Plus there's excel interop.  

What in the hell is sporklisp?
=============================

It's a lisp interpreter I wrote in VBA, with Excel interop via the ability to call user defined
functions from excel.  A lot of the concepts came from lisp wizards, particulalry the
Structure and Interpretation of Computer Programs (SICP), Peter Norvig's excellent tutorials
for LisPy and JScheme, and Christian Queinnec's Lisp in Small Pieces.

Why is it called sporklisp?
=============================

Originally, I intended to build a small lisp for a VBA-based set of libraries called
SPORK (Spoon's Operations Research Kit).  These libs were designed to support a sizeable
discrete event simulation, and ended up including a ton of generic libraries and features
grafted onto VBA to facilitate a more functional style of programming.  sporklisp was
a piece of SPORK, but its existence merited a stand-alone library.

In the name of all that is pure and good in this world, why VBA?
================================================================

In short: environment, education, and experimentation.

I come from an Operations Research background.
In my little corner of the world, they preach that VBA is the
coin of the realm.  OR types use spreadsheets a lot, VBA is readily available.
It's almost like saying PHP is the common language of the web....They're both weak
scripting languages that have a surprising amount of market penetration
due to just "being everywhere".  Prior to discovering the One True Way, I spent
a LOT of time in VBA, using it as an environment to explore computer science and
software engineering, while facilitating my OR job.  Eventually, I outgrew the bounds
of VBA, although I was mightily impressed to see "just" how far you can push VBA if you're
so inclined.  sporklisp is an example of pushing VBA to do things it really wasn't intended to.
It's also a way to provide an embedded programming environment for a lisp dialect that can be
used in office products.  Finally, having lisp in the spreadsheet can provide some advantages
although I do most of my "real" work in a lisp repl (these days it's Clojure).  Finally,
most lispers go through a learning phase where, as in the Structure and Interpretation of Computer
Programs (by Abelson and Sussman), they show you how to implement your own lisp.  Lisp in Small Pieces
by Queinnec also explores a lot of ways to implement your favorite lisp dialect.  By the time I started
getting decent in Clojure, Common Lisp, and Scheme, I felt of sufficient maturity to implement a lisp.
This is my first foray into doing so.  It's been great fun and incredibly educational.  I highly recommend
you try it sometime :)

How did you do it?
==================

I've spent unholy amounts of time in VBA over the last 5 years,
and toward the end of my last project (prior to porting a 42K LOC simulation library +
generic data structures + graph library + tons-of-stuff-vba-is-missing + my discrete
event simulation) I decided to roll my own lisp in VBA.  At the time, I was already
moving to clojure anyway, and had re-organized the initial OOP-designed simulation
structure (yes, VBA can do decent OOP quite well without some of the masochism in
other languages - also without many of the benefits) into something with a much
more functional style (to facilitate porting to a functional programming language).
My idea was to use a little Lisp dialect for scripting entity behaviors and other
lightweight stuff (I wanted first class functions too dammit).  I used Peter Norvig's
LisPy example in Python to get me started (which is a lightweight implementation, but
it uses a lot of Python's features to easily import primitives...and python has first
class functions already).  I had to cross-reference the implementation from Norvig's
JScheme (which LisPy was a subset of), and adapt the solution to the unique challenges
in VBA.  Finally, Chapter 4 of SICP (Structure and Interpretation of Computer Programs)
was invaluable, as was Lisp in Small Pieces.

What kind of a lisp is sporklisp?  Which cows are sacred?
=========================================================

I was working from scheme-based source material, and I have been developing - A LOT -
in Clojure over the past year.  So, sporklisp is a lisp-1 (functions and vars share the
same namespace). I spliced together some things I really liked from Clojure:
the generic sequence library from Clojure, as well as lazy sequences.  So...sporklisp is
currently a bit of a scheme/clojure hybrid.  It has lexical scope.  It has clojure-style
literals for vectors, maps, and sets (currently VBA arrays and dictionaries, respectively).
It also has reader support for JSON literals, which may be useful.  It does NOT have
persistent data structures that make clojure cool, so youre stuck with the humble list.
It DOES have lazy/infinite sequences, so you can play some cool tricks.  The sequence library
alone is probably worth using overe stock VBA.  Keywords exist, as per Clojure.  I dont have them
mirroring clojure semantics exactly, so you can't "apply" them to a map to get the associated value...
yet.  I am close to having macros implemented, at which point I might re-write a ton of sporklisp in
sporklisp.

What about performance, and stability?
======================================

sporklisp is actually fast enough for lots of scripting...although I havent really tested the upper
bounds of its reasonable performance (i.e. doing numeric work).  I would not recommend building a monster
spreadsheet and having tons of lispfuncs in the formulas...that might be untenable (I don't know yet).
sporklisp currently has some ineffeciencies - the largest of which is VBA.  VBA is an interpreted language,
and it has a really wonky garbage collector.  Large programs, or large spreadsheets, may cause sporklisp to
leak memory, or they may just be slow.  On the plus side, sporklisp's evaluation model uses VBA dictionaries to
store environments, so cost of looking up definitions is pretty cheap.  On the down side, the base type for
lists is actually the vba collection, which is a bit heavy weight.  Most lisps use a simple cons-cell, or a pair,
or represent lists as chunks of arrays, which provide much better cache coherence and speed.  Even in VBA, itd be
better for sporklisp to use variant arrays for everything, where it's currently using collections.  Still, sporklisp
is fast enough for you, old man!

Can I use sporklisp from Excel?
===============================

The cool thing (and something I'm still working on in my spare time) is just using lisp
for your formula logic in excel.  I need to do some more stuff to make the interop sweeter
- like wrapping array formulas into operations like map or other functions that return
multiple values.  Still, you can eval lisp expressions as formulas and chain them along
in a reactive manner just like normal excel stuff.  You can even bind vars from excel
and make them available in the lisp environment.  Plenty of stuff doesn't work, but
there's a pretty sizeable core language implemented already.  Macros are not implemented,
although they're a skip and a jump away.  Lots of stuff is implemented in pure VBA,
for "speed" purposes (that seems silly since VBA is interpreted anyway).

If I "really" wanted a reactive spreadsheet, Id just wrap one in Swing with clojure
driving the dataflow evaluation.  Much easier (and it's been done already), and you
basically just get another way to talk to the repl.  This is more of a toy project, but
it's sliding up the cool scale because, as it's VBA, it's embeddable in any Office product
(as long as MS continues to support VBA).  Just like clojure lets me live in Java without
writing Java, one could theoretically live in VBA without writing VBA under sporklisp.

-Tom Spoon
