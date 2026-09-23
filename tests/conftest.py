import shutil
import subprocess

import pytest
from PIL import Image

from kinolist.models import Film


@pytest.fixture
def film() -> Film:
    return Film(
        title="Терминатор",
        year=1984,
        rating="8.0",
        countries=["США", "Великобритания"],
        description="Киборг из будущего прибывает в 1984 год.",
        directors=["Джеймс Кэмерон"],
        actors=["Арнольд Шварценеггер", "Линда Хэмилтон", "Майкл Бин", "Пол Уинфилд"],
        poster=Image.new("RGB", (360, 540), "gray"),
        kp_id=507,
        genres=["фантастика", "боевик", "триллер"],
        main_genre="фантастика",
    )


@pytest.fixture
def mp4_file(tmp_path):
    """Минимальный mp4-файл, созданный ffmpeg. Тест пропускается, если ffmpeg недоступен."""
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg is None:
        pytest.skip("ffmpeg не найден")
    path = tmp_path / "Terminator.mp4"
    subprocess.run(
        [
            ffmpeg,
            "-loglevel",
            "error",
            "-y",
            "-f",
            "lavfi",
            "-i",
            "color=c=black:s=16x16:d=0.2",
            "-pix_fmt",
            "yuv420p",
            str(path),
        ],
        check=True,
    )
    return str(path)
