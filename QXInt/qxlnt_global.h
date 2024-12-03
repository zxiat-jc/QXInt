#pragma once

#include <QtCore/qglobal.h>

#ifndef BUILD_STATIC
# if defined(QXINT_LIB)
#  define QXINT_EXPORT Q_DECL_EXPORT
# else
#  define QXINT_EXPORT Q_DECL_IMPORT
# endif
#else
# define QXINT_EXPORT
#endif
