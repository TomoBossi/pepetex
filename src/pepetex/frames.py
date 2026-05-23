import tempfile
import argparse
from pathlib import Path
from collections.abc import Iterator
from contextlib import contextmanager

import cv2

import utils
import errors

RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout11.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_name}"/>
</Relationships>
"""
SLIDE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:mv="urn:schemas-microsoft-com:mac:vml"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:pvml="urn:schemas-microsoft-com:office:powerpoint"
    xmlns:com="http://schemas.openxmlformats.org/drawingml/2006/compatibility"
    xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
    xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"
    xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor">
    <p:cSld>
        <p:bg>
            <p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2">
                        <a:alphaModFix/>
                    </a:blip>
                    <a:stretch>
                        <a:fillRect/>
                    </a:stretch>
                </a:blipFill>
            </p:bgPr>
        </p:bg>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="53" name="Shape 53"/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>
"""

def frame_slide_xml_iterator(start_image_number: int, end_image_number: int) -> Iterator[dict[str, str]]:
    """
    Returns an iterator of {"slide": "<slide_xml>", "slide_rels": "<slide_rels_xml>"} dicts
    where each dict represents single frame .jpg image background slide and its rels.
    """
    for i in range(start_image_number, end_image_number + 1):
        yield {
            "slide": SLIDE_XML,
            "slide_rels": RELS_XML.format(**{"image_name": f"image{i}.jpg"})
        }

@contextmanager
def video_capture_iterator(video_path: Path):
    """
    Yields frames of the video pointed at by video_path and handles clean exit.
    """
    capture = cv2.VideoCapture(video_path)
    if not capture.isOpened():
        capture.release()
        raise RuntimeError(f"video file '{video_path}' could not be opened")
    try:
        yield capture
    finally:
        capture.release()

def frames_directory(
    pptx_directory_path: Path,
    video_path: Path,
    slide_number: int,
    frame_step: int = 1,
    frame_size: float = 1.0,
    frame_quality: float = 0.75
) -> None:
    """
    Inserts .jpg frames of the video pointed at video_path as slides,
    starting at position slide_number, of the extracted .pptx file directory 
    pointed at by pptx_directory_path.
    """
    utils.set_image_content_types(pptx_directory_path)
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        encoding_parameters = [cv2.IMWRITE_JPEG_QUALITY, max(1, min(100, int(frame_quality*100)))]
        with video_capture_iterator(video_path) as capture:
            frame_number = 0
            returned = True
            while returned:
                returned, frame = capture.read()
                if frame_number % frame_step == 0:
                    if frame_size != 1.0:
                        height, width, _ = frame.shape
                        scaled_height, scaled_width = max(1, int(height*frame_size)), max(1, int(width*frame_size))
                        frame = cv2.resize(frame, (scaled_width, scaled_height), interpolation=cv2.INTER_AREA)
                    cv2.imwrite(tmp_dir_path / f"{frame_number}.jpg", frame, encoding_parameters)
                frame_number += 1
        start_image_number, end_image_number = utils.add_images(pptx_directory_path, tmp_dir_path)
        slide_xmls =  frame_slide_xml_iterator(start_image_number, end_image_number)
        utils.insert_slides(pptx_directory_path, slide_xmls, end_image_number - start_image_number + 1, slide_number)

def frames(
    pptx_path: Path,
    video_path: Path,
    slide_number: int | None = None,
    frame_step: int = 1,
    frame_size: float = 1.0,
    frame_quality: float = 0.75
) -> None:
    """
    Inserts frames of video at video_path as slides starting at position slide_number
    of the .pptx file or extracted .pptx file directory pointed at by pptx_path.
    """
    if slide_number is None:
        slide_number = 1 + utils.get_slide_count(pptx_path)
    return utils.pptx_path_handler(pptx_path, frames_directory, [video_path, slide_number, frame_step, frame_size, frame_quality])

def main():
    parser = argparse.ArgumentParser(description="Inserts frames of a video as individual slides starting at a given slide number.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-vp", "--video-path", type=str, required=True, help="Path to the video file to extract frames from.")
    parser.add_argument("-s", "--slide-number", type=int, required=True, help="Starting slide number at which video frames will be inserted as new slides (counting from 1).")
    parser.add_argument("-fs", "--frame-step", type=int, default=1, help="Number of slides that separate contiguous extracted video frames. 1 means no frames are skipped.")
    parser.add_argument("-fsz", "--frame-size", type=float, default=1.0, help="Frame size scaling factor, used for shrinking frames relative to the video resolution. It must be greater than 0 and lower or equal to 1.")
    parser.add_argument("-fq", "--frame-quality", type=float, default=0.75, help="Frame .jpg quality. It must be greater than 0 (0%% quality) and lower or equal to 1 (100%% quality).")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_video_path = Path(args.video_path)
    arg_slide_number = args.slide_number
    arg_frame_step = args.frame_step
    arg_frame_size = args.frame_size
    arg_frame_quality = args.frame_quality
    errors.error_validation_path_missing(arg_pptx_path)
    errors.error_validation_file_missing(arg_video_path)
    errors.error_validation_slide_numbers_out_of_range([arg_slide_number], utils.get_slide_count(arg_pptx_path) + 1)
    errors.error_validation_negative_numbers({"-fs (--frame-step)": arg_frame_step, "-fsz (--frame-size)": arg_frame_size, "-fq (--frame-quality)": arg_frame_quality})
    errors.error_validation_greater_than({"-fq (--frame-quality)": arg_frame_quality, "-fsz (--frame-size)": arg_frame_size}, 1.0)
    frames(arg_pptx_path, arg_video_path, arg_slide_number, arg_frame_step, arg_frame_size, arg_frame_quality)

if __name__ == "__main__":
    main()
