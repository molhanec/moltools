# Assorted various utility functions. Public domain

__all__ = ['get_page', 'safe_save']

from urllib import splithost
import httplib
import os


def do_get(url, headers={}):
    host, path = splithost(url[len('http:'): ])
    con = httplib.HTTPConnection(host)
    con.request("GET", path, headers=headers)
    res = con.getresponse()
    return res


def get_page(url, referer = None, max_redirect = 10):
    """
    Tries to download page referred by url using GET method.
    It follows these HTTP redirects (max_redirect times maximally):
        301 Moved Permanently
        302 Found
        303 See Other
        305 Use Proxy
        307 Temporary Redirect
    If you don't specify referer, the url is also used as referer.    
    """
    if referer is None:
        referer = url
    res = do_get(url, {'Referer': referer})
    data = res.read()
    if res.status in (httplib.MOVED_PERMANENTLY, httplib.FOUND, httplib.SEE_OTHER, httplib.USE_PROXY, httplib.TEMPORARY_REDIRECT):
        if max_redirect > 0:
            headers = dict(res.getheaders())
            return get_page(headers['location'], referer, max_redirect - 1)
        return None
    return data


def safe_save(data, filename, ext = '', dir = '.', compare_content = False):
    """
    Saves data into the filenam.ext (or filename if ext is empty) in
    the directory dir (current dir by default).
    If file with given name it will save the content into the file with
    filename (2).ext name. If it exists than filename (3).ext etc.
    If ext is empty it won't add dot to filename.
    If ext starts with dot, the dot is removed, so
    ext='bmp' and ext='.bmp' leads to same result.
    (Only first one is stripped.)
    If compare_content is true than if file with given filename
    (i.e. filename.ext, filename (2).ext etc.)
    already exists than it is read and compared with data. If they
    are the same the function exits returning False.
    Returns true if the file was written.
    Function expects that no other process changes content of the
    directory. If it does the file existence detection can fail
    and existing file can be overwritten.
    """
    new_filename = os.path.join(dir, filename)
    if ext <> '':
        if ext[0] == '.':  # strip dot from extension
            ext = ext[1:] 
        new_filename += '.%s' % ext
    i = 2
    while os.path.exists(new_filename):
        if compare_content:
            with open(new_filename, 'rb') as f:
                content = f.read()
            if content == data:
                return False
        new_filename = os.path.join(dir, '%s (%i)' % (filename, i))
        if ext <> '': new_filename += '.%s' % ext
        i += 1
    with open(new_filename, 'wb') as f:
        f.write(data)     
    return True
