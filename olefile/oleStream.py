# --- OleStream ---------------------------------------------------------------

class OleStream(io.BytesIO):
    """
    OLE2 Stream

    Returns a read-only file object which can be used to read
    the contents of a OLE stream (instance of the BytesIO class).
    To open a stream, use the openstream method in the OleFileIO class.

    This function can be used with either ordinary streams,
    or ministreams, depending on the offset, sectorsize, and
    fat table arguments.

    Attributes:

        - size: actual size of data stream, after it was opened.
    """
    # FIXME: should store the list of sects obtained by following
    # the fat chain, and load new sectors on demand instead of
    # loading it all in one go.

    def __init__(self, fp, sect, size, offset, sectorsize, fat, filesize, olefileio):
        """
        Constructor for OleStream class.

        :param fp: file object, the OLE container or the MiniFAT stream
        :param sect: sector index of first sector in the stream
        :param size: total size of the stream
        :param offset: offset in bytes for the first FAT or MiniFAT sector
        :param sectorsize: size of one sector
        :param fat: array/list of sector indexes (FAT or MiniFAT)
        :param filesize: size of OLE file (for debugging)
        :param olefileio: OleFileIO object containing this stream
        :returns: a BytesIO instance containing the OLE stream
        """
        log.debug('OleStream.__init__:')
        log.debug('  sect=%d (%X), size=%d, offset=%d, sectorsize=%d, len(fat)=%d, fp=%s'
            %(sect,sect,size,offset,sectorsize,len(fat), repr(fp)))
        self.ole = olefileio
        # this check is necessary, otherwise when attempting to open a stream
        # from a closed OleFileIO, a stream of size zero is returned without
        # raising an exception. (see issue #81)
        if self.ole.fp.closed:
            raise OSError('Attempting to open a stream from a closed OLE File')
        # [PL] To detect malformed documents with FAT loops, we compute the
        # expected number of sectors in the stream:
        unknown_size = False
        if size == UNKNOWN_SIZE:
            # this is the case when called from OleFileIO._open(), and stream
            # size is not known in advance (for example when reading the
            # Directory stream). Then we can only guess maximum size:
            size = len(fat)*sectorsize
            # and we keep a record that size was unknown:
            unknown_size = True
            log.debug('  stream with UNKNOWN SIZE')
        nb_sectors = (size + (sectorsize-1)) // sectorsize
        log.debug('nb_sectors = %d' % nb_sectors)
        # This number should (at least) be less than the total number of
        # sectors in the given FAT:
        if nb_sectors > len(fat):
            self.ole._raise_defect(DEFECT_INCORRECT, 'malformed OLE document, stream too large')
        # optimization(?): data is first a list of strings, and join() is called
        # at the end to concatenate all in one string.
        # (this may not be really useful with recent Python versions)
        data = []
        # if size is zero, then first sector index should be ENDOFCHAIN:
        if size == 0 and sect != ENDOFCHAIN:
            log.debug('size == 0 and sect != ENDOFCHAIN:')
            self.ole._raise_defect(DEFECT_INCORRECT, 'incorrect OLE sector index for empty stream')
        # [PL] A fixed-length for loop is used instead of an undefined while
        # loop to avoid DoS attacks:
        for i in range(nb_sectors):
            log.debug('Reading stream sector[%d] = %Xh' % (i, sect))
            # Sector index may be ENDOFCHAIN, but only if size was unknown
            if sect == ENDOFCHAIN:
                if unknown_size:
                    log.debug('Reached ENDOFCHAIN sector for stream with unknown size')
                    break
                else:
                    # else this means that the stream is smaller than declared:
                    log.debug('sect=ENDOFCHAIN before expected size')
                    self.ole._raise_defect(DEFECT_INCORRECT, 'incomplete OLE stream')
            # sector index should be within FAT:
            if sect<0 or sect>=len(fat):
                log.debug('sect=%d (%X) / len(fat)=%d' % (sect, sect, len(fat)))
                log.debug('i=%d / nb_sectors=%d' %(i, nb_sectors))
##                tmp_data = b"".join(data)
##                f = open('test_debug.bin', 'wb')
##                f.write(tmp_data)
##                f.close()
##                log.debug('data read so far: %d bytes' % len(tmp_data))
                self.ole._raise_defect(DEFECT_INCORRECT, 'incorrect OLE FAT, sector index out of range')
                # stop reading here if the exception is ignored:
                break
            # TODO: merge this code with OleFileIO.getsect() ?
            # TODO: check if this works with 4K sectors:
            try:
                fp.seek(offset + sectorsize * sect)
            except Exception:
                log.debug('sect=%d, seek=%d, filesize=%d' %
                    (sect, offset+sectorsize*sect, filesize))
                self.ole._raise_defect(DEFECT_INCORRECT, 'OLE sector index out of range')
                # stop reading here if the exception is ignored:
                break
            sector_data = fp.read(sectorsize)
            # [PL] check if there was enough data:
            # Note: if sector is the last of the file, sometimes it is not a
            # complete sector (of 512 or 4K), so we may read less than
            # sectorsize.
            if len(sector_data)!=sectorsize and sect!=(len(fat)-1):
                log.debug('sect=%d / len(fat)=%d, seek=%d / filesize=%d, len read=%d' %
                    (sect, len(fat), offset+sectorsize*sect, filesize, len(sector_data)))
                log.debug('seek+len(read)=%d' % (offset+sectorsize*sect+len(sector_data)))
                self.ole._raise_defect(DEFECT_INCORRECT, 'incomplete OLE sector')
            data.append(sector_data)
            # jump to next sector in the FAT:
            try:
                sect = fat[sect] & 0xFFFFFFFF  # JYTHON-WORKAROUND
            except IndexError:
                # [PL] if pointer is out of the FAT an exception is raised
                self.ole._raise_defect(DEFECT_INCORRECT, 'incorrect OLE FAT, sector index out of range')
                # stop reading here if the exception is ignored:
                break
        # [PL] Last sector should be a "end of chain" marker:
        # if sect != ENDOFCHAIN:
        #     raise IOError('incorrect last sector index in OLE stream')
        data = b"".join(data)
        # Data is truncated to the actual stream size:
        if len(data) >= size:
            log.debug('Read data of length %d, truncated to stream size %d' % (len(data), size))
            data = data[:size]
            # actual stream size is stored for future use:
            self.size = size
        elif unknown_size:
            # actual stream size was not known, now we know the size of read
            # data:
            log.debug('Read data of length %d, the stream size was unknown' % len(data))
            self.size = len(data)
        else:
            # read data is less than expected:
            log.debug('Read data of length %d, less than expected stream size %d' % (len(data), size))
            # TODO: provide details in exception message
            self.size = len(data)
            self.ole._raise_defect(DEFECT_INCORRECT, 'OLE stream size is less than declared')
        # when all data is read in memory, BytesIO constructor is called
        io.BytesIO.__init__(self, data)
        # Then the OleStream object can be used as a read-only file object.
