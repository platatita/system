### link: https://discussions.apple.com/thread/2115778?tstart=0

### commandes:
1. brew install ffmpeg
2. ffmpeg -i video.mp4 -i sub.srt -c:v copy -c:a copy -c:s mov_text -metadata:s:s:0 language=eng out.mp

#### description:
- video.mp4 is your source video clip.
- sub.srt is your subrip subtitle file.
