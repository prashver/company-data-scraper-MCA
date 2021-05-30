""" A command-line parser you won't hate
"""
import shlex
import collections
from itertools import tee
from functools import update_wrapper


class Action(object):
    """ Action module class
    """
    shlex = shlex
    namedtuple = collections.namedtuple
    OrderedDict = collections.OrderedDict

    class Option(object):
        """ Abstract base for flags and options
        """

        def __init__(
            self,
            short=None, long=None,
            *, type=None
        ):
            if (
                (short and not isinstance(short, str)) or
                (long and not isinstance(long, str))
            ):
                raise TypeError('option name should be a string')

            if not long and short and len(short) > 1:
                long = short
                short = None
            if short and len(short) != 1:
                raise TypeError(
                    'short form should be at most of one character')

            self.type = type
            self.long = long
            self.short = short

        def __call__(self, old=None, new=None):
            """ Called on each occurrence of this option

                old -- current value obtained from
                       a previous call to this method;
                       might be None
                new -- string taken from the command line;
                       None if self.type is None

                Shall return a new value that could be passed
                to the action function
            """
            raise NotImplementedError(
                'Option.__call__ should be overridden'
                ' by a descendant')

        def __repr__(self):
            class_name = self.__class__.__name__
            attributes = []
            if self.short:
                attributes.append(repr(self.short))
            if self.long:
                attributes.append(repr(self.long))
            if self.type:
                attributes.append('type=' + repr(self.type))
            return '{class_name}({attributes})'.format(
                class_name=class_name,
                attributes=', '.join(attributes))

    class Flag(Option):
        """ Like `-q`, `--quiet`
        """

        def __call__(self, old, new):
            return True

    class Count(Option):
        """ Like `-v`, `--verbose`
        """

        def __call__(self, old, new):
            return (old if old else 0) + 1

    class Key(Option):
        """ Like `-c32`, `--count=32` """

        def __init__(
            self,
            short=None, long=None,
            *, type=str
        ):
            super().__init__(short, long, type=type)

        def __call__(self, old, new):
            if old is not None:
                raise RuntimeError(
                    'key {} should be specified at most once'.format(
                        self.long or self.short))
            return self.type.__call__(new)

    def __init__(self):
        # actions are created with `@action` decorator;
        # during execution, we shall select one action
        # depending on what was passed through command line
        self.actions = {}
        self.default_action = None

    def __call__(self, function):
        """ Make a function into an action
            and record it as such
        """
        name = function.__code__.co_name
        self.actions[name] = self._make_action(function)

        return self.actions[name]

    def context(self):
        """ Create separate action parser
        """
        return type(self)()

    def default(self, function):
        """ Decorator for a default action
            which is invoked when no other action
            is specified
        """
        if self.default_action is not None:
            raise TypeError('there could be at most one default action')

        self.default_action = self._make_action(function)
        return self.default_action

    def execute(self, argv):
        """ Act as per arguments
        """
        shlex = type(self).shlex

        if type(argv) is str:
            argv = shlex.split(argv)
        if type(argv) is not list:
            raise TypeError('argv should be a list')

        action = None
        first_positional = None
        positionals = filter(lambda x: not x.startswith('-'), argv)
        for positional in positionals:
            first_positional = positional
            break
        if first_positional in self.actions:
            argv.remove(first_positional)
            action = self.actions[first_positional]
        elif self.default_action is not None:
            action = self.default_action
        elif first_positional and len(first_positional) > 0:
            raise RuntimeError(
                'no such action: `{}`'.format(first_positional))
        else:
            raise RuntimeError('no action specified')

        args, leftover = self._parse_command_line(action, argv)
        if leftover and not action.is_variadic:
            raise TypeError('too many arguments')

        return action.__call__(*leftover, **args)

    def _make_action(self, function):
        """ Iterate through function type annotations
            and accordingly fill in fields `arguments` and `options`

            Raise `TypeError` in case of unacceptable annotations
        """
        if function.__code__.co_name.startswith('<'):
            raise TypeError(
                'an action should only be'
                ' a function defined through `def`')

        options = dict()
        arguments = self.OrderedDict()

        def wrapper(*args, **keywords):
            return function(*args, **keywords)
        action = update_wrapper(wrapper, function)

        code = function.__code__
        args = code.co_varnames[:code.co_argcount]
        kwargs = code.co_varnames[-code.co_kwonlyargcount:]
        annotations = function.__annotations__

        for argname in args:
            # positional arguments are taken from function arguments
            annotation = annotations.get(argname, str)
            if not callable(annotation):
                raise TypeError('annotation should be callable')
            arguments[argname] = annotation

        for optname in kwargs:
            # options are derived from function kw-only arguments
            annotation = annotations.get(optname)
            mapper = self._normalize_annotation(optname, annotation)
            options[optname] = mapper

        action.options = options
        action.arguments = arguments
        action.is_variadic = len(code.co_varnames) > code.co_argcount

        return action

    def _parse_command_line(self, action, argv):
        """ Make arguments dict for supplied action

            Return pair of
                arguments dict for action
                and unconsumed args
        """
        opts, positionals = self._parse_options(action, argv)
        args, leftover = self._parse_arguments(action, positionals)

        call = {}
        call.update(opts)
        call.update(args)
        return call, leftover

    def _parse_options(self, action, argv):
        """ Return a pair whose
            first element is populated dict of options,
            and the second element is part of argv left untouched
        """
        dash_dash_index = None
        try:
            dash_dash_index = argv.index('--')
        except ValueError:
            dash_dash_index = len(argv)
        optargv = argv[:dash_dash_index]
        options = {}
        unconsumed = []
        positional_only = argv[dash_dash_index:]

        # record first arg as unconsumed
        # if it is not an option,
        # because in the loop below
        # we record unconsumed only from lookahead,
        # and the first element never gets to the lookahead
        if len(optargv) > 0:
            first_arg = optargv[0]
            if not self._is_option(first_arg):
                unconsumed.append(first_arg)

        # that `None` shall be a pair to the last element
        optargv += [None]
        firsts, nexts = tee(optargv)
        next(nexts, None)
        pairs = zip(firsts, nexts)
        for arg, next_arg in pairs:
            next_arg_used = False
            if self._is_long_option(arg):
                next_arg_used = self._consume_long_option(
                    action, arg, next_arg, options, unconsumed)
            elif self._is_option(arg):
                next_arg_used = self._consume_short_option(
                    action, arg, next_arg, options, unconsumed)

            if (
                next_arg is not None and
                not self._is_option(next_arg) and
                not next_arg_used
            ):
                # case like `-h xyz`,
                # where current option is saturated or takes no args
                unconsumed.append(next_arg)

        positionals = unconsumed + positional_only
        return (options, positionals)

    def _consume_short_option(
        self, action, arg, next_arg, opts, unconsumed
    ):
        """ Modify `opts` to contain an option from `arg`
            Return True if `next_arg` was used up, False otherwise
        """
        key = arg[1]  # `arg[0]` is the dash
        remains = arg[2:]

        name = None
        mapper = None
        for each_name, each_mapper in action.options.items():
            if each_mapper.short == key:
                name, mapper = each_name, each_mapper
                break
        else:
            unconsumed.append(arg)
            return False

        old = opts.get(name)
        if mapper.type is not None:
            if remains:
                # `-c32`, remains are '32'
                argvalue = mapper.type.__call__(remains)
                opts[name] = mapper.__call__(old, argvalue)
                return False
            elif next_arg is not None:
                # `-c 32`, remains are '', next argument is '32'
                argvalue = mapper.type.__call__(next_arg)
                opts[name] = mapper.__call__(old, argvalue)
                return True
            else:
                raise TypeError(
                    'option `-{}` needs an argument'.format(key))
        elif remains:
            # case like `-vvvvvvc32`
            opts[name] = mapper.__call__(old, None)
            return self._consume_short_option(
                action, '-' + remains, next_arg, opts, unconsumed)

        opts[name] = mapper.__call__(old, None)
        return False

    def _consume_long_option(
        self, action, arg, next_arg, opts, unconsumed
    ):
        """ Modify `opts` to contain a long option from `arg`
            Return True if `next_arg` was used up, False otherwise
        """
        arg = arg[len('--'):]
        key, *value = arg.split('=', 1)

        name = None
        mapper = None
        for this_name, this_mapper in action.options.items():
            if this_mapper.long == key:
                name, mapper = this_name, this_mapper
                break
        else:
            unconsumed.append(arg)
            return False

        old = opts.get(name)

        if mapper.type is not None:
            if value:
                [value] = value
                argvalue = mapper.type.__call__(value)
                opts[name] = mapper.__call__(old, argvalue)
                return False
            elif next_arg is not None:
                argvalue = mapper.type.__call__(next_arg)
                opts[name] = mapper.__call__(old, argvalue)
                return True
            else:
                raise TypeError(
                    'option `--{}` requires an argument'.format(key))
        elif value:
            raise TypeError(
                'option `--{}` does not take arguments'.format(key))

        opts[name] = mapper.__call__(old, None)
        return False

    def _parse_arguments(self, action, argv):
        """ Match `argv` to respective fields in `action.arguments`

            Return pair of
                arguments dict
                and list of unconsumed parts of command line
        """
        arguments = {}
        leftover = []

        dash_dash_index = None
        try:
            dash_dash_index = argv.index('--')
        except ValueError:
            dash_dash_index = len(argv)
        maybe_positionals = argv[:dash_dash_index]
        positional_only = argv[dash_dash_index + 1:]

        positionals = []
        for arg in maybe_positionals:
            (leftover
                if self._is_option(arg)
                else positionals).append(arg)
        positionals += positional_only
        leftover += positionals[len(action.arguments):]

        args = zip(positionals, action.arguments.items())
        for arg, (name, mapper) in args:
            arguments[name] = mapper.__call__(arg)

        return arguments, leftover

    def _normalize_annotation(self, name, annotation):
        """ Return a mapper appropriate to the annotation passed
            name -- argument name
            annotation -- its annotation, if exists
        """
        Key = self.Key
        Flag = self.Flag
        Option = self.Option

        if annotation is None:
            return Key(name[0], name)

        elif type(annotation) is type:
            long = name.replace('_', '-')
            short = name[0]
            if issubclass(annotation, Option):
                # notation `verbose: action.Count`
                return annotation(short, long)
            elif issubclass(annotation, bool):
                # notation `quiet: bool`
                return Flag(short, long)
            else:
                # notation `depth: int`
                return Key(short, long, type=annotation)

        elif type(annotation) is tuple and len(annotation) == 3:
            # notation `reset_hard: ('r', 'hard', bool)`
            short, long, mapper = annotation
            return Key(short, long, type=mapper)

        if isinstance(annotation, Option):
            # notation `follow_symlinks: action.Flag('n', 'follow')`
            return annotation

        raise TypeError(
            'option annotation should be'
            ' an instance of action.Option,'
            ' a callable, or a triple (str, str, callable)')

    @staticmethod
    def _is_long_option(s):
        """ True iff a string starts with a double dash
            and has something more in it
        """
        return len(s) > len('--') and s.startswith('--')

    @staticmethod
    def _is_option(s):
        """ True iff a string starts with a dash
            and has something more in it
        """
        return len(s) > len('-') and s.startswith('-')


import sys  # nopep8
sys.modules[__name__] = Action()
