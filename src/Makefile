# Makefile for spdocx
# Copyright (C) 2021 ~ 2025 drangon <drangon_zhou (at) hotmail.com>

# This program is free software: you can redistribute it and/or modify it
# under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program; if not, see <http://www.gnu.org/licenses/>.

TARGET = libspdocx.so

SRCS += $(wildcard *.cpp)

OBJS = $(foreach src, $(SRCS), $(basename $(src)).o)
DEPS = $(foreach src, $(SRCS), $(basename $(src)).zzdep)

GXX = g++

CFLAGS += -g -O2 -Wall -pipe -DPIC -fPIC
LDFLAGS += -lpugixml -lzip 

.PHONY : all clean

all : $(TARGET)

$(TARGET) : $(OBJS)
	$(GXX) $(CFLAGS) -shared -o $@ $(OBJS) $(LDFLAGS)
	mkdir -p ../bin
	cp $(TARGET) ../bin

clean :
	rm -f $(TARGET) *.o *.zzdep

.SUFFIXES : .cpp

.cpp.o :
	$(GXX) $(CFLAGS) -MMD -MF $*.zzdep -c -o $@ $<

-include $(DEPS)

