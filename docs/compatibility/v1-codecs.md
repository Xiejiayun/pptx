# v1 codec compatibility matrix

| Capability | PowerPoint 2010+ | PowerPoint current | Keynote current | LibreOffice current | Google Slides import |
| --- | --- | --- | --- | --- | --- |
| Master/Layout/Theme preserve/edit | native | native | preserved / native basics | native basics | preserved / imported |
| Linear gradient | native | native | native | native | native with possible approximation |
| Path gradient | native | native | native basics | native basics | degraded diagnostic |
| Theme color + alpha transforms | native | native | minor rendering variance diagnostic | native basics | imported |
| Embedded MP3/WAV | native when client codec exists | native when client codec exists | client codec dependent | client codec dependent | import dependent |
| Embedded MP4 | native when H.264/AAC is available | native | client codec dependent | client codec dependent | import dependent |
| External media | portability warning | portability warning | preserved | preserved | link/import dependent |
| Click-to-play action | native | native | client dependent | native basics | imported behavior may change |
| Autoplay/loop/volume preferences | timing codec metadata | timing codec metadata | diagnostic | diagnostic | diagnostic |

库验证 OOXML 容器、关系和 content type，不承诺媒体编解码器可用性。未知扩展默认保留；任何降级都产生 diagnostic，不静默删除或栅格化。

