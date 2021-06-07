# Ballot Operations Output Tool & Help (BOOTH)

BOOTH (Ballot Operations Output Tool & Help) is a Microsoft Excel Add-In. BOOTH
helps users process, analyze, and visualize precinct voting data through the
voting process (check-in, completing the ballot, and ballot submission).

The latest *stable* release is
[BOOTH v3.0](https://personal.egr.uri.edu/macht/BOOTH_Voting_Package_V3.0.zip),
and more information (including help documentation) for that version can be
found [here](https://web.uri.edu/urivotes/tools/booth/). This repository
contains code for the in-progress version 4.0, with support for more log file
types, improvements to timers, and a port of the entire codebase to C# (as a
VSTO addin) from VBA.

## Overview of Features

The latest revision of BOOTH (the `master` branch) includes:

- Support for processing log files of four different ballot marking/scanning
devies and one electronic pollbook:
    - DS200
    - VSAP BMD
    - Dominion Imagecast X
    - Dominion Imagecast Evolution
    - PollPad (EPB) (*Not implemented yet*)
- Summary statistics for log files of the supported ballot marking/scannig
  devices
- Support for processing an entire folder of logs at once, without importing
  them into Excel first.
- A timers tool for timing election processes, incuding:
    - Voter Arrival Timer
    - Ballot Scanning Timer
    - BMD Timer
    - Voter Check-In Timer
    - Throughput Timer
    - Voting Booth Timer
