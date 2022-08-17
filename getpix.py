#!/usr/bin/env python3

import os, io, re, configparser
import imaplib, email
import pprint
from PIL import Image, ExifTags, ImageOps
from datetime import datetime

config = configparser.ConfigParser(allow_no_value=True,interpolation=configparser.ExtendedInterpolation())

# Mailbox settings
imapHost = 'hosting.axonsolutions.com'
imapUser = 'pix@eric.and.dans.wedding'
imapPass = '9@iwLwXx5T'

# Minimum dimensions for attached photo
imgMinWidth = 480
imgMinHeight = 480

# Images larger than this will be resized
imgMaxWidth = 2000
imgMaxHeight = 1000

# Where to store image files
imgHome = '/var/www/clients/client1/web7/home/obriend1/'
imgOriginals = imgHome + 'Pictures/'
imgStore = '/var/www/clients/client1/web7/web/wp-content/folder-slider/'

def set_default_config():
    global config

    config.add_section('imap']
    config.set('imap', 'server',          'localhost')
    config.set('imap', 'username',        'email@example.com')
    config.set('imap', 'password',        'password')
    config.set('imap', 'inboxFolder',     'Inbox')
    config.set('imap', 'processedFolder', 'Processed')
    config.set('imap', 'skippedFolder',   'Skipped')

    config.add_section('paths']
    config.set('paths', 'basePath',   '/tmp/')
    config.set('paths', 'originals',  '${basePath}originals/')
    config.set('paths', 'processed',  '${basePath}processed/')

    config.add_section('images']
    config.set('images', 'minWidth',   '480')
    config.set('images', 'minHeight',  '480')
    config.set('images', 'maxWidth',  '2000')
    config.set('images', 'maxHeight', '1000')

    config.add_section('security']
    config.set('security', 'rolesAllowed', 'contributor, author, editor, administrator')

    config.add_section('authorizedEmails']

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
    return True

def do_setup():
    global i
    os.makedirs(config['paths']['originals'], mode=0o755, exist_ok=True)
    os.makedirs(config['paths']['processed'], mode=0o755, exist_ok=True)
    create_mailbox(i, config['imap']['processedFolder'])
    create_mailbox(i, config['imap']['skippedFolder'])
    create_mailbox(i, config['imap']['errorsFolder'])

def path_munge(fn, d, email):
    eFromLocalPart, eFromDomain = eFromAddr.split('@')
    p = '/'.join( [ eFromDomain, eFromLocalPart, d + "-" + fn ] )
    return p


set_default_config()
config.read('getpix.ini')

i = imaplib.IMAP4_SSL( config['imap']['server'] )
i.login(config['imap']['username'], config['imap']['password'])
do_setup()

i.select( config['imap']['inboxFolder'] )
tmp, msgList = i.search(None, 'UNDELETED')

for num in msgList[0].split():
    print('Processing message number', num)

    # Retrieve and parse an email
    data = i.fetch(num, '(RFC822)')
    e = email.message_from_bytes( data[1][0][1] )
    eSubj = e.__getitem__('Subject')
    eDate = e.__getitem__('Date')
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
                if imgWidth >= config['images']['minWidth'] and imgHeight >= config['images']['minHeight']:

                    # Reduce image to max size, if necessary
                    factorX = imgWidth / imgMaxWidth
                    factorY = imgHeight / imgMaxHeight
                    if ( factorX ) > 1  or ( factorY ) > 1:
                        if factorX > factorY:
                            img = img.resize( (imgMaxWidth, int(imgHeight/factorX)), resample=None, box=None, reducing_gap=None )
                        else:
                            img = img.resize( (int(imgWidth/factorY), imgMaxHeight), resample=None, box=None, reducing_gap=None )

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
                    img.save( config['paths'['processed'] + imgFn )

                    # Save the email
                    i.append( config['imap']['processedFolder'], None, None, e.as_bytes() )

                else:
                    i.append( config['imap']['skippedFolder'], None, None, e.as_bytes() )

            n = n + 1
            i.store( num, '+FLAGS', '\\Deleted' )
            print()

        # End for msgPart in e.walk()
    # End if not e.is_multipart()

i.close()

#    print('Message: {0}\n'.format(num))
#    print('Message %s\n%s\n' % (num, data[0][1].decode("utf-8")))
#    pprint.pprint(data[0][1])
#    break
