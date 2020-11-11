#! /bin/bash -l

# Install commands used within the test file
sudo apt install sysbench stress-ng p7zip-full s-tui -y

# To enable graphical mode use 's-tui'

echo
echo "* Running cpu the benchmark: '7z b'"
time 7z b

# The sample of the output:
# 7-Zip [64] 16.02 : Copyright (c) 1999-2016 Igor Pavlov : 2016-05-21
# p7zip Version 16.02 (locale=en_US.UTF-8,Utf16=on,HugeFiles=on,64 bits,16 CPUs Intel(R) Core(TM) i7-10700K CPU @ 3.80GHz (A0655),ASM,AES-NI)
#
# Intel(R) Core(TM) i7-10700K CPU @ 3.80GHz (A0655)
# CPU Freq: - - - - - - - - -
#
# RAM size:   64243 MB,  # CPU hardware threads:  16
# RAM usage:   3530 MB,  # Benchmark threads:     16
#
#                        Compressing  |                  Decompressing
# Dict     Speed Usage    R/U Rating  |      Speed Usage    R/U Rating
#          KiB/s     %   MIPS   MIPS  |      KiB/s     %   MIPS   MIPS
#
# 22:      58040  1262   4475  56462  |     639199  1567   3479  54517
# 23:      55947  1263   4512  57003  |     628103  1555   3495  54344
# 24:      54559  1298   4520  58663  |     613472  1540   3497  53847
# 25:      53777  1343   4571  61401  |     616691  1573   3489  54883
# ----------------------------------  | ------------------------------
# Avr:            1292   4519  58382  |             1559   3490  54398
# Tot:            1425   4005  56390


echo
echo "* Running cpu the stress test for 1 minute: 'sysbench cpu --time=60 --threads=32 run'"
time sysbench cpu --time=60 --threads=32 run

# The sample of the output:
# CPU speed:
#     events per second: 20347.29

# General statistics:
#     total time:                          60.0011s
#     total number of events:              1220878

# Latency (ms):
#          min:                                    0.64
#          avg:                                    1.57
#          max:                                   57.31
#          95th percentile:                        3.82
#          sum:                              1919338.68

# Threads fairness:
#     events (avg/stddev):           38152.4375/1220.39
#     execution time (avg/stddev):   59.9793/0.02

echo
echo "* Running cpu the stress test for 1 minute: 'stress-ng -t 60 --cpu 16 --vm 16 --vm-bytes 4G --fork 16'"
time stress-ng -t 60 --cpu 16 --vm 16 --vm-bytes 4G --fork 16

# The sample of the output:
# stress-ng: info:  [317186] dispatching hogs: 16 cpu, 16 vm, 16 fork
# stress-ng: info:  [317186] successful run completed in 60.17s (1 min, 0.17 secs)
