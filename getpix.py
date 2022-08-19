#!/usr/bin/env python3

import sys, os, io, re, configparser
import imaplib, email
import pprint
from PIL import Image, ExifTags, ImageOps
from datetime import datetime

config = configparser.ConfigParser(allow_no_value=True,interpolation=configparser.ExtendedInterpolation())

class msgPart:
    def __init__(self, p):
        self.raw = p.as_string()
        self.contentType = str( p.get_content_type() )
        self.contentDisp = str( p.get_content_disposition() )
        self.fileName = None
        self.fileData = None
        if self.contentDisp in [ "inline", "attachment" ]:
            self.fileName = str( p.get_filename() ).replace('.jpeg', '.jpg')
            self.fileData = p.get_payload(decode=True)

    def saveFileData(self, path, fn):
        os.makedirs(path, mode=0o755, exist_ok=True)
        msgFile = open( path + "/" + fn, "xb" )
        msgFile.write( self.fileData )
        msgFile.close()


class message:
    def __init__(self, i, num):

        # Retrieve and parse an email
        data = i.fetch(num, '(UID RFC822)')
        e = email.message_from_bytes( data[1][0][1] )

        self.raw = e.as_string()
        self.isMultipart = e.is_multipart()
        self.messageId = e.__getitem__('Message-Id')
        self.fromName, self.fromAddr = email.utils.parseaddr( e.get('From') )
        self.fromLocalPart, self.fromDomain = self.fromAddr.split('@')
        self.subject = e.__getitem__('Subject')
        self.date = e.__getitem__('Date')
        self.dateTime=datetime.strptime(self.date, '%a, %d %b %Y %H:%M:%S %z')
        self.compactDate = self.dateTime.strftime('%Y%m%d%H%M%S')
        self.msgParts = []

        if self.isMultipart:
            n = 0
            for p in e.walk():
                self.msgParts.append( msgPart(p) )
                print("Making multipart list", n, type(self.msgParts[n]))
                n += 1

def wr_log(f, msg):
    now = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    return True

def get_config():
    global config

    config.add_section('imap')
    config.set('imap', 'server',          'localhost')
    config.set('imap', 'username',        'email@example.com')
    config.set('imap', 'password',        'password')
    config.set('imap', 'inboxFolder',     'Inbox')
    config.set('imap', 'processedFolder', 'Processed')
    config.set('imap', 'skippedFolder',   'Skipped')

    config.add_section('paths')
    config.set('paths', 'basePath',   '/tmp/')
    config.set('paths', 'originals',  '${basePath}originals/')
    config.set('paths', 'processed',  '${basePath}processed/')

    config.add_section('images')
    config.set('images', 'minWidth',   '480')
    config.set('images', 'minHeight',  '480')
    config.set('images', 'maxWidth',  '2000')
    config.set('images', 'maxHeight', '1000')

    config.add_section('security')
    config.set('security', 'rolesAllowed', '')

    config.add_section('authorizedEmails')
    config.add_section('bannedEmails')

    config.read('getpix.ini')

    return True


def mailbox_exists(i, mailbox):
    x, y = i.list()
    if x == 'OK':
        for mb in y:
            if mb.split()[2].decode() == mailbox:
                return True
    return False

def create_mailbox(i, mailbox):
    if not mailbox_exists(i, mailbox):
        x, y = i.create(mailbox)
        if x == 'OK':
            return True
        else:
            return False
    else:
        return True

def is_authorized_email(e):
    rolesAllowed = re.sub(' *, *', ',', config['security']['rolesAllowed']).split(',')
    if e in config['bannedEmails']:
        return False
    x = False
    y = e in config['authorizedEmails']
    return x or y

def do_setup():
    global i
    os.makedirs(config['paths']['originals'], mode=0o755, exist_ok=True)
    os.makedirs(config['paths']['processed'], mode=0o755, exist_ok=True)
    create_mailbox(i, config['imap']['processedFolder'])
    create_mailbox(i, config['imap']['skippedFolder'])
    create_mailbox(i, config['imap']['errorsFolder'])

def move_msg(i, e, f):
    return True

def path_munge(fn, d, email):
    eFromLocalPart, eFromDomain = eFromAddr.split('@')
    p = '/'.join( [ eFromDomain, eFromLocalPart, d + "-" + fn ] )
    return p


get_config()

i = imaplib.IMAP4_SSL( config['imap']['server'] )
try:
    i.login(config['imap']['username'], config['imap']['password'])
except:
    print('Failed to log in to IMAP server.')
    sys.exit(1)

do_setup()

i.select( config['imap']['inboxFolder'] )
tmp, msgList = i.search(None, 'UNDELETED')

for num in msgList[0].split():
    print('Processing message number', num)

    msg = message(i, num)

    if not ( msg.isMultipart and is_authorized_email(msg.fromAddr) ):

        print( '> Skipping' )
        i.append( config['imap']['skippedFolder'], None, None, msg.emailObj.as_bytes() )

    else:

        n = 0
        for p in msg.msgParts:

            if p.contentType in [ "image/jpeg", "image/png" ]:

                img = Image.open( io.BytesIO(p.fileData) )

                # Correct the orientation of the image
                img = ImageOps.exif_transpose(img)

                imgWidth, imgHeight = img.size

                # Skip images below size threshold
                if imgWidth >= config.getint('images', 'minWidth') and imgHeight >= config.getint('images', 'minHeight'):

                    # Reduce image to max size, if necessary
                    factorX = imgWidth / config.getint('images', 'maxWidth')
                    factorY = imgHeight / config.getint('images', 'maxHeight')
                    if ( factorX ) > 1  or ( factorY ) > 1:
                        if factorX > factorY:
                            img = img.resize( ( config.getint('images', 'maxWidth'), int(imgHeight/factorX) ), resample=None, box=None, reducing_gap=None )
                        else:
                            img = img.resize( ( int(imgWidth/factorY), config.getint('images', 'maxHeight') ), resample=None, box=None, reducing_gap=None )

                    # Save the original attachment
                    imgSeq = msg.compactDate + '{:03d}'.format(n)
                    filePath = msg.fromDomain + '/' + msg.fromLocalPart
                    filePath = config['paths']['originals'] + filePath
                    p.saveFileData(filePath, imgSeq + "-" + p.fileName)
                    print("Saving original:", filePath + "/" + imgSeq + "-" + p.fileName)

                    # Save the processed image
                    imgFn = imgSeq + '_' + re.sub('[`~@#$%^*{}[]<>/?]', "", msg.subject + '_(' + msg.fromName + ')')
                    imgFn = imgFn.replace(' ', '_').replace('__', '_')
                    imgFn = imgFn + '.' + str(p.fileName).split('.')[-1]
                    print("Saving processed", config['paths']['processed'] + imgFn)
                    img.save( config['paths']['processed'] + imgFn )

                    # Save the email
                    i.append( config['imap']['processedFolder'], None, None, msg.emailObj.as_bytes() )

                else:
                    i.append( config['imap']['skippedFolder'], None, None, msg.emailObj.as_bytes() )

            n = n + 1
#           i.store( num, '+FLAGS', '\\Deleted' )

        # End for p in msg.msgParts
    # End if not msg.isMultipart

i.close()
