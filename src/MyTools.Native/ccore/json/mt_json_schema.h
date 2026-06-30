#pragma once

#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

int mt_json_has_schema_version(const char* json, size_t size);

#ifdef __cplusplus
}
#endif
