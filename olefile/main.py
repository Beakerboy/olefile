# --------------------------------------------------------------------
# This script can be used to dump the directory of any OLE2 structured
# storage file.

def main():
    """
    Main function when olefile is runs as a script from the command line.
    This will open an OLE2 file and display its structure and properties
    :return: nothing
    """
    import sys, optparse

    DEFAULT_LOG_LEVEL = "warning" # Default log level
    LOG_LEVELS = {
        'debug':    logging.DEBUG,
        'info':     logging.INFO,
        'warning':  logging.WARNING,
        'error':    logging.ERROR,
        'critical': logging.CRITICAL
        }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)

    parser.add_option("-c", action="store_true", dest="check_streams",
        help='check all streams (for debugging purposes)')
    parser.add_option("-v", action="store_true", dest="extract_customvar",
        help='extract all document variables')
    parser.add_option("-p", action="store_true", dest="extract_customprop",
                      help='extract all user-defined propertires')
    parser.add_option("-d", action="store_true", dest="debug_mode",
        help='debug mode, shortcut for -l debug (displays a lot of debug information, for developers only)')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")

    (options, args) = parser.parse_args()

    print('olefile version {} {} - https://www.decalage.info/en/olefile\n'.format(__version__, __date__))

    # Print help if no arguments are passed
    if len(args) == 0:
        print(__doc__)
        parser.print_help()
        sys.exit()

    if options.debug_mode:
        options.loglevel = 'debug'

    # setup logging to the console
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], format='%(levelname)-8s %(message)s')

    # also enable the module's logger:
    enable_logging()

    for filename in args:
        try:
            ole = OleFileIO(filename)#, raise_defects=DEFECT_INCORRECT)
            print("-" * 68)
            print(filename)
            print("-" * 68)
            ole.dumpdirectory()
            for streamname in ole.listdir():
                if streamname[-1][0] == "\005":
                    print("%r: properties" % streamname)
                    try:
                        props = ole.getproperties(streamname, convert_time=True)
                        props = sorted(props.items())
                        for k, v in props:
                            # [PL]: avoid to display too large or binary values:
                            if isinstance(v, (basestring, bytes)):
                                if len(v) > 50:
                                    v = v[:50]
                            if isinstance(v, bytes):
                                # quick and dirty binary check:
                                for c in (1,2,3,4,5,6,7,11,12,14,15,16,17,18,19,20,
                                          21,22,23,24,25,26,27,28,29,30,31):
                                    if c in bytearray(v):
                                        v = '(binary data)'
                                        break
                            print("   ", k, v)
                    except Exception:
                        log.exception('Error while parsing property stream %r' % streamname)

                    try:
                        if options.extract_customprop:
                            variables = ole.get_userdefined_properties(streamname, convert_time=True)
                            if len(variables):
                                print("%r: user-defined properties" % streamname)
                                for index, variable in enumerate(variables):
                                    print('\t{} {}: {}'.format(index, variable['property_name'],variable['value']))

                    except:
                        log.exception('Error while parsing user-defined property stream %r' % streamname)
                elif options.extract_customvar and streamname[-1]=="WordDocument":
                    print("%r: document variables" % streamname)
                    variables = ole.get_document_variables()

                    for index, var in enumerate(variables):
                        print('\t{} {}: {}'.format(index, var['var_name'], var['value'][:50]))
                    print("")


            if options.check_streams:
                # Read all streams to check if there are errors:
                print('\nChecking streams...')
                for streamname in ole.listdir():
                    # print name using repr() to convert binary chars to \xNN:
                    print('-', repr('/'.join(streamname)),'-', end=' ')
                    st_type = ole.get_type(streamname)
                    if st_type == STGTY_STREAM:
                        print('size %d' % ole.get_size(streamname))
                        # just try to read stream in memory:
                        ole.openstream(streamname)
                    else:
                        print('NOT a stream : type=%d' % st_type)
                print()

            # for streamname in ole.listdir():
            #     # print name using repr() to convert binary chars to \xNN:
            #     print('-', repr('/'.join(streamname)),'-', end=' ')
            #     print(ole.getmtime(streamname))
            # print()

            print('Modification/Creation times of all directory entries:')
            for entry in ole.direntries:
                if entry is not None:
                    print('- {}: mtime={} ctime={}'.format(entry.name,
                        entry.getmtime(), entry.getctime()))
            print()

            # parse and display metadata:
            try:
                meta = ole.get_metadata()
                meta.dump()
            except Exception:
                log.exception('Error while parsing metadata')
            print()
            # [PL] Test a few new methods:
            root = ole.get_rootentry_name()
            print('Root entry name: "%s"' % root)
            if ole.exists('worddocument'):
                print("This is a Word document.")
                print("type of stream 'WordDocument':", ole.get_type('worddocument'))
                print("size :", ole.get_size('worddocument'))
                if ole.exists('macros/vba'):
                    print("This document may contain VBA macros.")

            # print parsing issues:
            print('\nNon-fatal issues raised during parsing:')
            if ole.parsing_issues:
                for exctype, msg in ole.parsing_issues:
                    print('- {}: {}'.format(exctype.__name__, msg))
            else:
                print('None')
            ole.close()
        except Exception:
            log.exception('Error while parsing file %r' % filename)


if __name__ == "__main__":
    main()

# this code was developed while listening to The Wedding Present "Sea Monsters"
