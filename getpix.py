#!/usr/bin/env python3

import sys, os, io, re, configparser
import imaplib, email
import pprint
from PIL import Image, ExifTags, ImageOps
from datetime import datetime

config = configparser.ConfigParser(allow_no_value=True,interpolation=configparser.ExtendedInterpolation())

class msgPart:
    def __init(self, p):
        self.msgPartObj = p
        self.contentType = str( p.get_content_type() )
        self.contentDisp = str( p.get_content_disposition() )
        self.fileName = str( p.get_filename() )
        self.fileData = p.get_payload(decode=True)

class message:
    def __init__(self, i, num):

        # Retrieve and parse an email
        data = i.fetch(num, '(UID RFC822)')

        self.emailObj = email.message_from_bytes( data[1][0][1] )
        self.isMultipart = self.emailObj.is_multipart()
        self.messageId = e.__getitem__('Message-Id')
        self.fromName, self.fromAddr = email.utils.parseaddr( e.get('From') )
        self.fromLocalPart, self.fromDomain = self.fromName.split('@')
        self.dubject = e.__getitem__('Subject')
        self.date = e.__getitem__('Date')
        self.dateTime=datetime.strptime(eDate, '%a, %d %b %Y %H:%M:%S %z')
        self.compactDate = eDateTime.strftime('%Y%m%d%H%M%S')
        self.msgParts = {}

        if self.isMultipart:
            n = 0
            for p in self.emailObj.walk():
                self.msgParts[n] = msgPart(p)
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

    # Retrieve and parse an email
    data = i.fetch(num, '(UID RFC822)')
    e = email.message_from_bytes( data[1][0][1] )
    eSubj = e.__getitem__('Subject')
    eDate = e.__getitem__('Date')
    eMsgId = e.__getitem__('Message-Id')
    eDateTime=datetime.strptime(eDate, '%a, %d %b %Y %H:%M:%S %z')
    eDate = eDateTime.strftime('%Y%m%d%H%M%S')
    eFromName, eFromAddr = email.utils.parseaddr( e.get('From') )
    eFromLocalPart, eFromDomain = eFromAddr.split('@')

    if not ( e.is_multipart() and is_authorized_email(eFromAddr) ):

        print( '> Skipping' )
        i.append( config['imap']['skippedFolder'], None, None, e.as_bytes() )

    else:

        n = 0
        for msgPart in e.walk():
            print( '> ContentType: ' + str( msgPart.get_content_type() ) )
            print( '> ContentDisposition: ' + str( msgPart.get_content_disposition() ) )
            print( '> FileName: ' + str( msgPart.get_filename() ) )

            if msgPart.get_content_type() in [ "image/jpeg", "image/png" ]:
                msgFilename = msgPart.get_filename().replace('.jpeg', '.jpg')

                # Create an image object from mail attachment
                msgFiledata = msgPart.get_payload(decode=True)
                img = Image.open( io.BytesIO(msgFiledata) )

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

                    print( '  > Format: ' + str(img.format) )
                    print( '  > Size: ', end='')
                    print( imgWidth, 'x', imgHeight )

                    # Save the original attachment
                    imgSeq = eDateTime.strftime('%Y%m%d%H%M%S') + '{:03d}'.format(n)
                    filePath = eFromDomain + '/' + eFromLocalPart + '/'
                    filePath = config['paths']['originals'] + filePath
                    os.makedirs(filePath, mode=0o755, exist_ok=True)
                    msgFile = open( filePath + imgSeq + '-' + str(msgFilename), "xb" )
                    msgFile.write( msgFiledata )
                    msgFile.close()

                    # Save the processed image
                    # name = YYYYMMDDnnn_${Name}_-_${Subject}.ext
                    imgFn = imgSeq + '_' + re.sub('[`~@#$%^*{}[]<>/?]', '', eSubj) + '_(' + eFromName.replace(' ', '_') + ')'
                    imgFn = imgFn.replace(' ', '_')
                    imgFn = imgFn + '.' + str(msgFilename).split('.')[-1]
                    print( imgStore, imgFn)
                    img.save( config['paths']['processed'] + imgFn )

                    # Save the email
                    i.append( config['imap']['processedFolder'], None, None, e.as_bytes() )

                else:
                    i.append( config['imap']['skippedFolder'], None, None, e.as_bytes() )

            n = n + 1
#           i.store( num, '+FLAGS', '\\Deleted' )
            print()

        # End for msgPart in e.walk()
    # End if not e.is_multipart()

i.close()

#    print('Message: {0}\n'.format(num))
#    print('Message %s\n%s\n' % (num, data[0][1].decode("utf-8")))
#    pprint.pprint(data[0][1])
#    break
