''''
Project 3
COMP 467
Brian Kelly
Due: 12/13/2023
'''

import sys
import argparse
import pymongo
import subprocess
import shlex
import xlsxwriter
from frameioclient import FrameioClient
from KEYS import FRAMES_TOKEN, DESTINATION_ID

def setup_mongodb():
    client = pymongo.MongoClient('localhost', 27017)  
    db = client['project2_db']  
    return db['metadata_collection'], db['content_collection']

def parse_arguments():
    parser = argparse.ArgumentParser(description="Project 3")
    parser.add_argument("--output", dest="output", help="Output xls")
    parser.add_argument("--process", dest="process", help="Process Video File")
    args = parser.parse_args()
    
    if args.process is None:
        print("Missing Video File")
        sys.exit(2)
    
    return args

def get_duration_from_ffmpeg(process_file):
    command = 'ffmpeg -i {} -hide_banner'.format(process_file)
    print(command)
    process = subprocess.Popen(shlex.split(command), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    for line in process.stdout.readlines():
        decoded = line.decode()
        if decoded.startswith('  Duration:'):
            return decoded.strip().split(",")[0].strip().split(' ')[1]
    return None

def calculate_total_frames(duration):
    timecode = duration.split(':')
    hour = int(timecode[0])
    min = int(timecode[1])
    second = float(timecode[2].split('.')[0])
    return int((second * 24) + (min * 60 * 24) + (hour * 60 * 60 * 24))

def find_work_within_video(collection, total_frames):
    withInVideo = []
    for work in collection.find():
        frame_range = work['frames'].split('-')
        if len(frame_range) == 2 and int(frame_range[1]) <= total_frames:
            withInVideo.append(work)
        elif len(frame_range) == 1 and int(frame_range[0]) <= total_frames:
            withInVideo.append(work)
    return withInVideo

def convert_time_code(frame):
    frames = frame % 24
    seconds = frame // 24
    mins = seconds // 60
    hours = mins // 60
    seconds %= 60
    mins %= 60
    hours %= 24
    frames = int((frames / 24.0) * 100) % 100
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(hours, mins, seconds, frames)

def find_time_code(frame_range):
    frame_range = frame_range.split('-')

    #Determine middle frame to use for image output for frame range
    if len(frame_range) == 2:
        frame = int((int(frame_range[0]) + (int(frame_range[1]))) / 2)
    #For single frame
    else:
        frame = int(frame_range[0])
    
    return convert_time_code(frame)

def generate_images_and_fill_xls(withInVideo, process_file, workbook_name='output.xlsx'):
    workbook = xlsxwriter.Workbook(workbook_name)
    worksheet = workbook.add_worksheet()

    for i, work in enumerate(withInVideo):
        timecode = find_time_code(work['frames'])
        command = 'ffmpeg -i {} -ss {} -vf \'scale=96:74:force_original_aspect_ratio=decrease\' -frames:v 1 output{}.png'.format(process_file, timecode[0:7], i)
        subprocess.call(shlex.split(command))
        image = "output{}.png".format(i)
        worksheet.write('A{}'.format(i + 1), work['location'])
        worksheet.write('B{}'.format(i + 1), work['frames'])
        worksheet.write('C{}'.format(i + 1), timecode)
        worksheet.insert_image('D{}'.format(i + 1), 'output{}.png'.format(i))
        upload_to_frameio(image)

    workbook.close()

def upload_to_frameio(image):
    client = FrameioClient(FRAMES_TOKEN)
    client.assets.upload(DESTINATION_ID, image)

def main():
    metadata_collection, content_collection = setup_mongodb()
    args = parse_arguments()
    
    duration = get_duration_from_ffmpeg(args.process)
    if duration is None:
        print("Could not get video duration.")
        sys.exit(1)

    total_frames = calculate_total_frames(duration)
    withInVideo = find_work_within_video(content_collection, total_frames)

    if args.output == 'xls':
        generate_images_and_fill_xls(withInVideo, args.process)


if __name__ == "__main__":
    main()











