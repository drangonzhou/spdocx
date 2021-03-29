// SPD_Def.h : spdocx common define
// Copyright (C) 2021 ~ 2021 drangon <drangon_zhou (at) hotmail.com>
//
// This program is free software: you can redistribute it and/or modify it
// under the terms of the GNU Lesser General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public License
// along with this program; if not, see <http://www.gnu.org/licenses/>.

#ifndef INCLUDED_SPD_DEF_H
#define INCLUDED_SPD_DEF_H

#define BEGIN_NS_SPD namespace spd {
#define END_NS_SPD }

#ifdef _WIN32
#ifdef LIBSPDOCX_EXPORTS
#define SPD_API __declspec(dllexport)
#else
#define SPD_API __declspec(dllimport)
#endif // BUILD_SPD_DLL
#else
#define SPD_API
#endif // _WIN32 else

BEGIN_NS_SPD
////////////////////////////////

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_DEF_H
