This library draws a parallel between command line args and function args in Python.
Positionals get passed to regular function arguments, and options or flags are mapped to keyword arguments.

For example, this command invocation::

    $ package install -u ffmpeg -v

Could be translated to this function call::

    package.install('ffmpeg', upgrade=True, verbose=1)

This library does the bridging automatically
using information supplied in form of decorators and type annotations.
To make a function accessible as a command-line action, decorate it with `action`::

    import sys
    import action

    @action
    def install(package_name, *, upgrade: action.Flag = False, verbose: action.Count = 0):
        """ Do the work
        """

    if __name__ == '__main__':
        sys.exit(action.execute(sys.argv[1:]))

All other exported symbols are described below.

----

@action
=======
The main decorator which is used to make actions from functions.
It takes a single function as an input and inspects its signature.

The name of the command being created
is drawn from the name of original function.

All arguments before splat are counted as positionals,
and those going after are options or flags.

Configuration through annotations
=================================
Client code could alter how certain arguments
are treated and presented by annotating its arguments.

One way to do so is to supply a constructor as an annotation::

    @action
    def add(x: int, y: int):
        print(x + y)

That constructor shall be called upon execution to coerce types
before passing arguments to the action invoked.

Positionals only support this kind of annotations.

Options, on the other hand, use callable annotations differently.
Each option or flag could occur many times,
therefore that behaviour should be covered by corresponding
annotation.
There are some sane defaults come already packaged.

Flag
----
Denotes whether some condition is truthy.
Could be specified any number of times on command line.
First occurrence sets the value to `True`.
Subsequent occurrences have no effect::

    @action
    def add(x: int, y: int, *, pad: action.Flag = False):
        result = x + y
        format = '{}'
        if pad:
            format = '{:04}'
        print(format.format(result))

Count
-----
Initially is `None`.
On first occurrence sets to one,
on each subsequent occurrence increments by one::

    @action
    def add(x: int, y: int, *, verbose: action.Count = 0):
        result = x + y

        if verbose > 3:
            print('augend:', x)
            print('addend:', y)
            print('sum:   ', result)
            print()
        elif verbose > 0:
            print('{} + {} = {}'.format(result))
        else:
            print(result)

Key
---
Generic value specified as a command-line option::

    @action
    def walk(*, depth: action.Key('depth', type=int)):
        ...

`Key` constructor has three arguments: `short`, `long` and `type`.
One of `short` or `long` is required.  `type` is `str` by default.

any callable
------------
There is also a shorthand notation for specifying a Key::

    @action
    def walk(*, depth: int):
        ...

Short and long names shall be deduced from the argument name.

(short, long, type) triple
--------------------------
Another shorthand for Key allows
to specify short and long names manually::

    @action
    def walk(*, depth: ('r', 'depth', int)):
        ...

Option abstract base
--------------------
On a low level, to know a value for an option, the command line
processor performs a folding operation over all occurrences
of a certain option.
Therefore, to have fine-grained control over the argument parsing
process, one could subclass `action.Option` to use it
instead of prepackaged annotations for options.
Subclass should override call method to take two arguments:
the old value and an option body.
That call method could either return a new value or throw an exception
to stop command line processing right away.
If call method returns a value, that value shall be passed
as old value on the next call.

@action.default
===============
The command line processor selects an action
whose name matches the first positional.
If there is no such action registered,
the command line processor attempts
to invoke the special action marked as default::

    @action.default
    @action
    def install(package):
        ...

    # `./prog.py install ffmpeg` shall invoke `install('ffmpeg')`
    # and `./prog.py ffmpeg` shall still invoke `install('ffmpeg')`

This decorator could also be used if the program
has a single action::

    @action.default
    def list_directory():
        ...

action.execute
==============
Look up a previously registered action whose name matches
first positional from command line,
match command-line arguments to selected action arguments
and invoke that action.

The first positional argument is hidden from the command invoked.

`action.execute` never calls `os.exit`,
so it could be used in an interactive prompt.

action.context
==============
If you want an isolated argument parser to avoid modification
of module-wide state, you could instantiate another `Action`
with this method.

Normally, an `Action` object is constructed in place
of `action` module when importing.

----

Coded with Love.


