#!/usr/bin/env bash
ps as | grep julius | grep -v grep | awk '{print $2}' | xargs kill
