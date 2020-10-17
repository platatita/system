#! /bin/bash -l

# Install commands used within the test file
sudo apt install mbw sysbench -y

echo
echo "* Running memory bandwidth test - 'mbw -t1 20000Gb'"
mbw -t1 20000Gb

# The sample of the output:
# Long uses 8 bytes. Allocating 2*2621440000 elements = 41943040000 bytes of memory.
# Getting down to business... Doing 10 runs per test.
# 0   Method: DUMB    Elapsed: 1.15262    MiB: 20000.00000    Copy: 17351.727 MiB/s
# 1   Method: DUMB    Elapsed: 1.15348    MiB: 20000.00000    Copy: 17338.896 MiB/s
# 2   Method: DUMB    Elapsed: 1.15484    MiB: 20000.00000    Copy: 17318.461 MiB/s
# 3   Method: DUMB    Elapsed: 1.15298    MiB: 20000.00000    Copy: 17346.294 MiB/s
# 4   Method: DUMB    Elapsed: 1.15282    MiB: 20000.00000    Copy: 17348.732 MiB/s
# 5   Method: DUMB    Elapsed: 1.15258    MiB: 20000.00000    Copy: 17352.390 MiB/s
# 6   Method: DUMB    Elapsed: 1.15191    MiB: 20000.00000    Copy: 17362.528 MiB/s
# 7   Method: DUMB    Elapsed: 1.15192    MiB: 20000.00000    Copy: 17362.287 MiB/s
# 8   Method: DUMB    Elapsed: 1.15329    MiB: 20000.00000    Copy: 17341.737 MiB/s
# 9   Method: DUMB    Elapsed: 1.15440    MiB: 20000.00000    Copy: 17325.017 MiB/s
# AVG Method: DUMB    Elapsed: 1.15308    MiB: 20000.00000    Copy: 17344.796 MiB/s


echo
echo "* Memory short details - 'sudo lshw -short -C memory'"
sudo lshw -short -C memory

# The sample of the output:
# H/W path         Device           Class          Description
# ============================================================
# /0/0                              memory         64KiB BIOS
# /0/39                             memory         64GiB System Memory
# /0/39/0                           memory         16GiB DIMM DDR4 Synchronous 4000 MHz (0,2 ns)
# /0/39/1                           memory         16GiB DIMM DDR4 Synchronous 4000 MHz (0,2 ns)
# /0/39/2                           memory         16GiB DIMM DDR4 Synchronous 4000 MHz (0,2 ns)
# /0/39/3                           memory         16GiB DIMM DDR4 Synchronous 4000 MHz (0,2 ns)
# /0/48                             memory         512KiB L1 cache
# /0/49                             memory         2MiB L2 cache
# /0/4a                             memory         16MiB L3 cache
# /0/100/14.2                       memory         RAM memory


echo
echo "* Running memory speed test - 'sysbench memory --memory-total-size=20000Gb --threads=16 run'"
sysbench memory --memory-total-size=20000Gb --threads=16 run

# The sample of the output:
# Total operations: 311096835 (31106333.30 per second)

# 303805.50 MiB transferred (30377.28 MiB/sec)

# General statistics:
#     total time:                          10.0001s
#     total number of events:              311096835

# Latency (ms):
#          min:                                    0.00
#          avg:                                    0.00
#          max:                                   16.00
#          95th percentile:                        0.00
#          sum:                               133112.28

# Threads fairness:
#     events (avg/stddev):           19443552.1875/896093.56
#     execution time (avg/stddev):   8.3195/0.08
