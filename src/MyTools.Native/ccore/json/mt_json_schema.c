#include "ccore/json/mt_json_schema.h"

int mt_json_has_schema_version(const char* json, size_t size) {
    static const char token[] = "\"schema_version\"";
    const size_t token_size = sizeof(token) - 1;

    if (json == 0 || size < token_size) {
        return 0;
    }

    for (size_t index = 0; index + token_size <= size; ++index) {
        size_t token_index = 0;
        while (token_index < token_size && json[index + token_index] == token[token_index]) {
            ++token_index;
        }
        if (token_index == token_size) {
            return 1;
        }
    }

    return 0;
}
