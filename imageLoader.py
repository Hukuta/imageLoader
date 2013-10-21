#!/usr/bin/env python
#-*- coding: UTF-8 -*-
"""Download pictures for the shop on the CMS php5shop"""

import os
import sys
import urllib2
import logging

if __name__ == '__main__':
    from xlrd import open_workbook
    from xlrd.biffh import XLRDError
    from Lang import Locale
    from PIL import Image
    # GUI gtk OR console mode
    GTK = (1 == len(sys.argv))
    if GTK:
        import gtk
        import gobject
        import threading

        # application error
        def error(message):
            md = gtk.MessageDialog(
                None,
                gtk.DIALOG_DESTROY_WITH_PARENT,
                gtk.MESSAGE_ERROR,
                gtk.BUTTONS_OK, message)
            md.run()
            md.destroy()
    else:
        # application error
        def error(message):
            logging.error(message)

    logging.basicConfig(format='%(levelname)s: %(message)s', level=logging.WARNING)

    # Getting list of images (prodId, url) from xls file
    def parse_xls(path):
        to_download = list()
        try:
            wb = open_workbook(path.decode())
        except IOError:
            error(Locale.NO_SUCH_FILE % path)
            return to_download
        except XLRDError, err:
            error(err.message)
            return to_download

        for s in wb.sheets():
            for row in range(s.nrows):
                try:
                    prod_id = str(int(s.cell(row, 0).value))
                except ValueError:
                    continue
                try:
                    images = str(s.cell(row, 5).value).split(' ')
                except IndexError:
                    continue
                if len(images):
                    counter = 0
                    for img in images:
                        if not counter:
                            to_download.append((prod_id, img.decode()))
                        else:
                            to_download.append(('_'.join((prod_id, str(counter))), img.decode()))
                        counter += 1
                        if counter > 255:
                            break
        if not len(to_download):
            error(Locale.NO_IMAGES)
        return to_download

    # HTTP request
    def request(uri):
        try:
            return urllib2.urlopen(uri, timeout=10).read()
        except (IOError, ValueError), ioErr:
            logging.warn(ioErr.message)
            return False

    def get_dir():
        folder = 'images'
        if not os.path.isdir(folder):
            os.mkdir(folder, 0777)
        folder4small = '/'.join((folder, 'small'))
        if not os.path.isdir(folder4small):
            os.mkdir(folder4small, 0777)
        if not os.path.isdir(folder4small):
            error(Locale.FOLDER_ERROR % folder4small)
            return False
        return folder

    badURLs = list()

    if GTK:
        window = gtk.Window()
        window.set_default_size(350, 200)
        window.set_position(gtk.WIN_POS_CENTER)
        window.set_title(Locale.WINDOW_TITLE)
        window.connect('destroy', lambda w: gtk.main_quit())

        vBox = gtk.VBox(True, 5)

        title = gtk.Label(Locale.ABOUT)
        vBox.add(title)

        progress = gtk.ProgressBar()

        vBox.add(progress)

        titlePrice = gtk.Label(Locale.SELECT_LABEL)
        vBox.add(titlePrice)

        downloadButton = gtk.Button(Locale.DOWNLOAD_BUTTON)
        downloadButton.set_data('path', '')

        stopButton = gtk.Button(Locale.STOP_BUTTON)

        def set_path(path):
            downloadButton.set_data('path', path)
            titlePrice.set_text(Locale.SELECTED_FILENAME % os.path.split(path)[-1:][0])

        def select():
            chooser_dialog = gtk.FileChooserDialog(
                title=Locale.SELECT_DIALOG_TITLE,
                parent=window,
                action=gtk.FILE_CHOOSER_ACTION_OPEN,
                buttons=(gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL, gtk.STOCK_OPEN, gtk.RESPONSE_OK)
            )
            file_filter = gtk.FileFilter()
            file_filter.set_name("xls")
            file_filter.add_pattern("*.xls")
            chooser_dialog.add_filter(file_filter)
            response = chooser_dialog.run()
            if response == gtk.RESPONSE_OK:
                set_path(chooser_dialog.get_filename())
            chooser_dialog.destroy()

        selectButton = gtk.Button(Locale.SELECT_BUTTON)
        selectButton.connect('clicked', lambda b: select())
        vBox.add(selectButton)

        checkbox = gtk.CheckButton(Locale.SKIPPING_CHECKBOX)

        stopEvent = threading.Event()
        stopped = threading.Event()

        stoppedDialogEvent = threading.Event()

        def stop_button_click(button):
            if stoppedDialogEvent.is_set():
                return
            stoppedDialogEvent.set()

            if button:
                logging.log(logging.DEBUG, 'stopButton Clicked')
                stopEvent.set()
            stopButton.hide()
            if button:
                logging.log(logging.DEBUG, 'waiting for stop daemon thread')
                stopped.wait()
                stopped.clear()
                logging.log(logging.DEBUG, 'daemon thread stopped')
            progress.hide()
            downloadButton.show()
            selectButton.show()
            checkbox.show()

            md = gtk.MessageDialog(
                window,
                gtk.DIALOG_DESTROY_WITH_PARENT,
                gtk.MESSAGE_INFO,
                gtk.BUTTONS_OK, Locale.DONE)
            md.run()
            md.destroy()

            errCount = len(badURLs)
            if 0 < errCount <= 20:
                error("\r\n".join((Locale.ERROR_URLS, "\r\n".join(badURLs))))
            elif errCount > 20:
                badURLsSlice = badURLs[:20]
                error("\r\n".join((Locale.ERROR_URLS, "\r\n".join(badURLsSlice), 'и др.')))
            stoppedDialogEvent.clear()

    def get_images(lst, folder, stop, stopped, badURLs, progressBar, skipExistFiles):
        if GTK:
            logging.log(logging.DEBUG, 'daemon thread started')
        i = 0
        c = len(lst)
        logging.log(logging.DEBUG, '%d images' % c)
        watermark = Image.open('watermark.png', 'r')
        for prodId, url in lst:
            if GTK and stop.is_set():
                progressBar.set_fraction(1.)
                break
            path2save = "".join((folder, '/', prodId, '.jpg'))
            path2saveSmall = "".join((folder, '/small/', prodId, '.jpg'))
            if skipExistFiles and os.path.isfile(path2save):
                #skipping already existing files
                logging.log(logging.DEBUG, 'skipping file %s' % path2save)
                i += 1
                continue
            logging.log(logging.DEBUG, 'request to %s' % url)
            contents = request(url)
            if not contents:
                #second attempt
                contents = request(url)
                if not contents:
                    badURLs.append(url)
                    i += 1
                    continue
            logging.log(logging.DEBUG, 'request is done')

            with open(path2save, 'wb') as fileHandle:
                fileHandle.write(contents)
            try:
                im = Image.open(path2save)
                im.thumbnail((500, 500), Image.ANTIALIAS)
                copyBig = im.copy()
                im.thumbnail((150, 150), Image.ANTIALIAS)
                im.save(path2saveSmall, "JPEG")
                copyBig.paste(watermark, (0, 0), watermark)
                copyBig.save(path2save, "JPEG")
            except StandardError, e:
                logging.log(logging.WARNING, e.message)
                badURLs.append(url)
                i += 1
                continue

            i += 1
            fraction = float(i) / c
            if GTK:
                progressBar.set_fraction(fraction)
            else:
                print int(fraction * 100), '%'
                #all finished
        if GTK:
            stopped.set()

    def download_list(list_to_download, folder):
        if GTK:
            progress.set_text(Locale.DOWNLOAD_PROCESS % len(list_to_download))
            stopped.clear()
            stopEvent.clear()
            logging.log(logging.DEBUG, 'creating daemon thread')
            thread = threading.Thread(target=get_images, args=(
                list_to_download, folder, stopEvent, stopped, badURLs, progress, checkbox.get_active()))
            thread.setDaemon(True)

            logging.log(logging.DEBUG, 'starting daemon thread')
            thread.start()

            def wait_while_done():
                thread.join(0.49)
                if stopped.is_set() or not thread.is_alive():
                    stop_button_click(None)
                else:
                    gobject.timeout_add(500, wait_while_done)

            wait_while_done()

        else:
            get_images(list_to_download, folder, None, None, badURLs, None,
                       len(sys.argv) == 3 and sys.argv[2] == '-skip')

    if GTK:

        def download_init(button):
            path = button.get_data('path')
            if '' == path:
                error(Locale.FILE_IS_NOT_SELECTED)
                return
            progress.set_text(Locale.PARSING_XLS)
            progress.set_fraction(0)
            progress.show()
            stopButton.show()
            checkbox.hide()
            downloadButton.hide()
            selectButton.hide()
            directory = get_dir()
            toDownloadList = parse_xls(path)
            if directory and len(toDownloadList):
                download_list(toDownloadList, directory)

        vBox.add(checkbox)

        downloadButton.connect('clicked', download_init)
        vBox.add(downloadButton)

        stopButton.connect('clicked', stop_button_click)
        vBox.add(stopButton)

        window.add(vBox)
        window.show_all()

        progress.hide()
        stopButton.hide()

        try:
            window.set_icon_from_file("icon.gif")
        except Exception, e:
            print e.message

        window.present()
        gtk.main()

    else:
        print '-' * 80
        print 'usage: python imageLoader.py path_to_file.xls [-skip]\r\n'
        print 'path_to_file.xls\t : your xls price for CMS php5shop'
        print '-skip\t : to skip existing images'
        print '-' * 80
        directory = get_dir()
        if directory:
            download_list(parse_xls(sys.argv[1]), directory)
            if len(badURLs):
                print Locale.ERROR_URLS
                for url in badURLs:
                    print url