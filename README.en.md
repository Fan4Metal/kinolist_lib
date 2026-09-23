# Kinolist Lib

[Русская версия](README.md)

Kinolist Lib is a Windows command-line tool (`kl`) that builds film lists in docx format and writes film metadata into mp4 tags. Film data is retrieved from the unofficial Kinopoisk API (kinopoiskapiunofficial.tech).

## Features

- Film lists in docx format with posters (A4 and A5 templates) or as a plain numbered list.
- Lists built from film titles given on the command line, from a text file, or from mp4 files in a directory.
- Writing a full film card (title, year, rating, countries, description, directors, actors, poster, genres, Kinopoisk id) into mp4 tags.
- Building lists offline from previously written tags, including files referenced by Windows shortcuts (`.lnk`).
- Renaming files with torrent-style names into `Title (year).ext`.
- Caching of API responses for one hour.

## Requirements

- Windows 10 or later.
- Python 3.13 or later and [uv](https://docs.astral.sh/uv/) for running from source.
- An API token from [kinopoiskapiunofficial.tech](https://kinopoiskapiunofficial.tech).

## Installation

### Installer

The installer (`Kinolist_Lib <version> Setup.exe`) places the program into the user profile, adds it to `PATH` and registers Explorer context-menu commands: a list from tags for a directory (with and without a cover), tag writing and removal for an mp4 file, and a list from a txt file.

### From source

```
git clone <repository url>
cd kinolist_lib
uv sync
```

The `kl` command is then available via `uv run kl`.

## Configuration

The API token is read from the `config.py` file located next to the program:

```python
KINOPOISK_API_TOKEN = "your-token"
```

## Usage

| Command | Description |
|---|---|
| `kl -m "Terminator" "Terminator 2" KP~319` | List from the given titles; `KP~id` specifies a Kinopoisk id directly. |
| `kl -f movies.txt -o movies.docx` | List from titles in a text file, one per line. |
| `kl -l [dir]` | List from mp4 file names in a directory. |
| `kl --loc [dir]` | List built from mp4 tags only, without network access. |
| `kl -t [file or dir]` | Writes tags into an mp4 file or into all mp4 files in a directory. |
| `kl -t file.mp4 -kp 406` | Writes tags of the film with the given Kinopoisk id. |
| `kl --cleartags [file or dir]` | Removes all tags; `--confirm` asks for confirmation first. |
| `kl -r *.mp4` | Renames files after confirmation. |

Modifiers: `-o` output file, `-s` shortened descriptions, `--txtlist` additional txt list, `-nf` plain list format, `-g` genres in the list, `--a5` A5 template, `--cover [text]` cover page first (without text the title is taken from the directory or file name), `--cover-name` names the output file after the cover title, `--sort` file order for `--loc` (`date`, `date_r`, `datem`, `datem_r`, `name`, `name_r`), `--test` search without creating a list, `--nocache`, `--clearcache`, `--pause` waits for Enter before exiting (used by the context-menu commands).

Full reference: `kl --help`.

## Tag layout

Data is stored in standard mp4 atoms (`©nam`, `©day`, `©gen`, `desc`, `ldes`, `covr`) and in free-form atoms `----:com.apple.iTunes:*` (`DIRECTOR`, `Actors`, `kpra`, `countr`, `kpid`, `genre`). A rating in `kpra` that starts with `i` (for example `i6.7`) is interpreted as an IMDb rating. The layout is compatible with files tagged by earlier versions.

## Development

```
uv run pytest
uv run ruff check .
uv run ruff format .
```

Tag tests require `ffmpeg` on `PATH`; otherwise they are skipped.

## Build

```
uv run python tools/make_release.py
```

The script builds the `dist\kl` directory with PyInstaller and, when Inno Setup 6 is installed, the installer `dist\Kinolist_Lib <version> Setup.exe`. The `--no-installer` option limits the build to the directory. The version is taken from `kinolist/__init__.py` and stamped into both the exe version resource and the installer.

## License

MIT, see [LICENSE](LICENSE).
