#pragma once

#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef struct mt_result {
    int code;
    const char* message;
} mt_result;

typedef struct mt_buffer {
    unsigned char* data;
    size_t size;
} mt_buffer;

#ifdef __cplusplus
}
#endif
