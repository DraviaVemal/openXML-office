cmake_minimum_required(VERSION 3.15)

project(ProtoBuf)

message("Proto Buffer build Start")

# Step 1: Prepare the Rust structs
message("Proto Buffer build Rust Proto Preparation")
# Step 2: Prepare the C# Class
message("Proto Buffer build C# Proto Preparation")
# Step 3: Prepare the Go interface
message("Proto Buffer build Go Proto Preparation")
# Step 4: Prepare the Java Class
message("Proto Buffer build Java Proto Preparation")
# Step 5: Copy schema's to targets
message("Proto Buffer copy results to target directory")

message("Proto Buffer build Complete")