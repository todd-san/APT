# encoding: utf-8
"""
Simple program used for updating KTEK.AssemblePattern(HpcJob) status' and stages as they solve on kendaeng02

Args:
    meta.JSON: file that should be auto-generated by KTEK HpcJob models
    status: char-string that will be sent to update the model in KTEK
    stage: (optional) char-string that will be sent to update the model in KTEK

"""
# imports
import os
import sys
import requests
import logging
import json
import time
from PIL import Image
# import urllib.parse


class KpatResults:
    def __init__(self, **kwargs):
        # hosted path example!
        # http://kendaeng02:8888/media/DAS/analysis/pr1/ph1/d1/s1/pattern/1_6.bmp

        # file path example!
        # /DAS/analysis/pr1/ph1/d1/s1/pattern/1_6.bmp

        # meta.json path example!
        # /share/analysis/pr1/ph1/d1/s1/pattern/

        self.auth = ('admin', 'kenda000')
        self.error = {'GET': list(), 'PATCH': list()}

        self.domain = None
        self.job_url = None
        self.results_url = None
        self.path = None
        self.pitch_width = None
        self.zone_image_width = None
        self.zone_image_height = None

        if kwargs:
            for key, value in kwargs.items():
                setattr(self, key, value)

    def api_patch(self, data):
        """
        try to patch results into restful endpoint
        """
        api_patch_url = ''.join(['http://', self.domain, self.results_url])

        try:
            r = requests.patch(url=api_patch_url, auth=self.auth, data=data)

            if str(r.status_code)[0] == '4' or str(r.status_code)[0] == '5':
                self.error['PATCH'].append("Http PATCH Failed: {}, {}".format(r.status_code, api_patch_url))
                raise requests.exceptions.HTTPError

        except requests.exceptions.HTTPError:
            log.error("Http PATCH Error: {}".format(self.error['PATCH']))
        except requests.exceptions.ConnectionError as err_c:
            self.error['PATCH'].append(str(err_c))
            log.error("Http PATCH Connection Error: {}".format(err_c))
        except requests.exceptions.Timeout as err_t:
            self.error['PATCH'].append(str(err_t))
            log.error("Http PATCH Timeout Error: {}".format(err_t))
        except requests.exceptions.RequestException as err_r:
            self.error['PATCH'].append(str(err_r))
            log.error("Http PATCH Request Error: {}".format(err_r))
        except Exception as e:
            self.error['PATCH'].append(str(e))
            log.error('Http PATCH Unhandled Error: {}'.format(e))

    def api_get(self):
        """
        try to grab job restful endpoint from ktek api
        :return: dict(job)
        """
        api_get_url = ''.join(['http://', self.domain, self.job_url])
        api_get = None

        try:
            api_get = requests.get(url=api_get_url, auth=self.auth)
            # log.info('Http GET CO: {}'.format(api_get.status_code))

            if str(api_get.status_code)[0] == '4' or str(api_get.status_code)[0] == '5':
                self.error['GET'] .append("Http GET Failed: {}, {}".format(api_get.status_code, api_get_url))
                raise requests.exceptions.HTTPError

        except requests.exceptions.HTTPError:
            log.error("Http GET Error: {}".format(self.error['GET']))
        except requests.exceptions.ConnectionError as err_c:
            self.error['GET'].append(str(err_c))
            log.error("Http GET Connection Error: {}".format(err_c))
        except requests.exceptions.Timeout as err_t:
            self.error['GET'].append(str(err_t))
            log.error("Http GET Timeout Error: {}".format(err_t))
        except requests.exceptions.RequestException as err_r:
            self.error['GET'].append(str(err_r))
            log.error("Http GET Request Error: {}".format(err_r))
        except Exception as e:
            self.error['GET'].append(str(e))
            log.error('Http GET Unhandled Error: {}'.format(e))

        if api_get:
            # log.info(api_get.text)
            return dict(json.loads(api_get.text))
            # return None
        else:
            return None

    def commit_results(self):
        # results for api endpoint
        api_results = {
            'tread': self.tread,
            'void': self.void,
            'zones': self.zones,
            'clean_imgs': self.clean_imgs,
            'orig_imgs': self.orig_imgs,
            'pitch_width': self.pitch_width,
            'zone_image_width': self.zone_image_width,
            'zone_image_height': self.zone_image_height,
        }
        self.api_patch(data=api_results)

        # pickup status, job_id from 'self.api_get()'
        # status, results & id for File System using a loop
        # write results out
        api_job = self.api_get()
        if api_job:
            file_results = {
                'treadpattern': {
                    'fields': {}
                },
                'patternassemble': {
                    'fields': {
                        'job_id': api_job['job_id'],
                        'status': api_job['status']
                    }
                }
            }

            for key in api_results:
                file_results['treadpattern']['fields'][key] = api_results[key]

            with open("results.json", 'w') as outfile:
                json.dump(file_results, outfile, indent=2)

            log.info('Results File Written')
        else:
            log.error('Http GET Failed: No Results File Written')
            pass

    @property
    def tread(self):
        found = False
        partial = "_Fulltread_0_crop"
        directory = os.listdir(os.getcwd())

        for item in directory:
            # log.debug('     {0}'.format(item))
            if partial in str(item):
                # log.debug('MATCH Found: {}'.format(item))
                found = True
                return self.set_path(item)

        if not found:
            return None

    @property
    def void(self):
        found = False
        partial = "_Void_auto"
        directory = os.listdir(os.getcwd())

        for item in directory:
            # log.debug('     {0}'.format(item))
            if partial in str(item):
                # log.debug('MATCH Found: {}'.format(item))
                found = True
                return self.set_path(item)

        if not found:
            return None

    @property
    def zones(self):
        found = False
        partial = "_Cropped"
        directory = os.listdir(os.getcwd())

        for item in directory:
            # log.debug('     {0}'.format(item))
            if partial in str(item):
                found = True
                return self.set_path(item)

        if not found:
            return None

    @property
    def orig_imgs(self):
        imgs = []
        partial = '.bmp'
        directory = os.listdir(os.getcwd())

        for item in directory:
            if partial in str(item):
                imgs.append(self.set_path(item))
        return ', '.join(imgs)

    @property
    def clean_imgs(self):
        imgs = []
        partial = '_clean.png'
        directory = os.listdir(os.getcwd())

        for item in directory:
            if partial in str(item):
                imgs.append(self.set_path(item))
        return ', '.join(imgs)

    def debug_log(self):
        attrs = [a for a in dir(self) if not a.startswith('__') and not callable(getattr(self, a))]
        for attr in attrs:
            log.debug('Status.{0}: {1}'.format(attr, getattr(self, attr)))

    def set_path(self, item):
        path = self.path.replace('/share', 'DAS')
        url = ['http://', '/'.join(map(lambda x: str(x).rstrip('/'), [self.domain, 'media', path, item]))]
        return ''.join(url)


def main(meta_file, rtn_file):
    log.info('RECEIVED META FILE: {}'.format(meta_file))

    m_kwargs = read_meta(file=meta_file)
    m_kwargs['pitch_width'] = read_rtn(file=rtn_file)
    m_kwargs['zone_image_width'], m_kwargs['zone_image_height'] = read_cropped_size()

    kr = KpatResults(**m_kwargs)
    kr.debug_log()
    kr.commit_results()

    if kr.error:
        log.info('ANALYSIS_STATUS.PY: UPDATE EXITED WITH ERRORS: {}'.format(kr.error))
    else:
        log.info('ANALYSIS_STATUS.PY: UPDATE FINISHED WITHOUT ERRORS')

    log.info('END: {}'.format(str(time.strftime('%Y-%m-%d | %H:%M:%SZ', time.gmtime()))))
    log.info('========================================================================================================')
    log.info('------------------------------------- KPAT_JOB_COMPLETE.PY ---------------------------------------------')
    log.info('-------------------------------------         END          ---------------------------------------------')
    log.info('========================================================================================================')


def read_meta(file):
    meta_kwargs = {}
    key_list = ['results_url', 'job_url', 'domain', 'path']

    with open(file) as json_file:
        full_meta = json.load(json_file)

    for key in dict.keys(full_meta):
        if key in key_list:
            meta_kwargs[key] = full_meta[key]

    return meta_kwargs


def read_rtn(file):
    lines = []
    with open(file) as txt_file:
        for line in txt_file:
            lines.append(line)
            # log.info('LINE: {}'.format(line))

    pitch_width = lines[0].split(',')[0]
    log.info('Pitch Width Guess: {}'.format(pitch_width))
    return pitch_width


def read_cropped_size():
    partial = "_Cropped"
    path = None
    directory = os.listdir(os.getcwd())

    # log.info("directory listing: {}".format(directory))
    # log.info("current path: {}".format(os.getcwd()))

    for item in directory:
        # log.debug('     {0}'.format(item))
        if partial in str(item):
            found = True
            path = item
            log.info('Found Cropped File: {}'.format(path))

    if path:
        im = Image.open(path)
        width, height = im.size
        log.info('Img Size,  w: {}, h: {}'.format(width, height))

    else:
        width = None
        height = None
        log.error('Could not find cropped image')

    return [width, height]


def init_logging():
    filelog = logging.getLogger('status')
    filelog.setLevel(logging.INFO)

    formatter = logging.Formatter('%(name)s - %(funcName)15s - %(levelname)s - %(message)s')

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)

    fh = logging.FileHandler(filename='kpat_results.log', mode='a')
    fh.setLevel(logging.INFO)

    ch.setFormatter(formatter)
    fh.setFormatter(formatter)

    filelog.addHandler(fh)
    filelog.addHandler(ch)
    return filelog


if __name__ == '__main__':
    log = init_logging()
    log.info('========================================================================================================')
    log.info('------------------------------------- KPAT_JOB_COMPLETE.PY ---------------------------------------------')
    log.info('-------------------------------------        START         ---------------------------------------------')
    log.info('========================================================================================================')
    log.info('START: {}'.format(str(time.strftime('%Y-%m-%d | %H:%M:%SZ', time.gmtime()))))
    try:
        main(meta_file=sys.argv[1], rtn_file=sys.argv[2])

    except IndexError:
        log.error("No 'meta.json' file provided")
        log.info('PROGRAM TERMINATED WITH ERRORS - IMPROPER NUMBER OF INPUT ARGUMENTS')
        log.info('END: {}'.format(str(time.strftime('%Y-%m-%d | %H:%M:%SZ', time.gmtime()))))




