from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter
import sys
import build
import os
from pathlib import Path

def get_args(arg_list):
    del arg_list[0]
    try:
        parser = ArgumentParser(description='program description', formatter_class=RawDescriptionHelpFormatter, prog='program_name')
        parser.add_argument('-a', '--adversary', dest="adversary", default=None, help="-a --adversary {adversary name} (use -l to list)", required=False)
        parser.add_argument('-f', '--filetype', dest="filetype", default="doc",
                            help="-f --filetype doc | docm | flatxml", required=False)
        parser.add_argument('-e', '--extension', dest="extension", default="doc", required=False,
                            help="-e --extension doc | docm | rtf")
        parser.add_argument('-c', '--count', dest="count", default=1, help="-c --count {# of docs to create}", required=False)
        parser.add_argument('-l', '--listadversaries', dest="listadversaries", action='store_true', help="-l --listadversaries : list available adversaries and exits")
        parser.add_argument('-o', '--outdir', dest="outdir", default="Default", help="-o --outdir {path\\to\\outdir}")
        parser.add_argument('-d', '--debug', dest="debug", default=False, action='store_true', help="-d --debug : print debug statements and playbook for each document")
        args = parser.parse_args(arg_list)
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print('[-] fatal error parsing arguments, error=' + repr(e) + ". for help please user --help")
        raise
    return args


args = get_args(sys.argv)

if args.listadversaries == True:
    adversary_list = os.listdir(os.path.join(Path(__file__).parent, "adversary"))
    for item in adversary_list:
        if os.path.isdir(os.path.join(Path(__file__).parent, "adversary", item)) and "__" not in item:
            print(item)

if not args.listadversaries and args.adversary and args.outdir:

    if os.path.isdir(args.outdir) == False:
        print("\n[!] outdir is either not defined or is not a valid folder path\n")
        raise ValueError

    build.build_files(
        adversary=args.adversary,
        out_dir=args.outdir,
        count=int(args.count),
        filetype=args.filetype,
        extension=args.extension,
        debug=args.debug
    )

