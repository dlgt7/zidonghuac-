#pragma once
#include <string>
#include <vector>
#include <map>
#include <sstream>
#include <cctype>
#include <algorithm>

namespace json {

class Value {
public:
    enum Type { Null, Bool, Number, String, Array, Object };
    
    Type type = Null;
    bool bool_val = false;
    double num_val = 0;
    std::string str_val;
    std::vector<Value> arr_val;
    std::map<std::string, Value> obj_val;
    
    Value() = default;
    Value(bool b) : type(Bool), bool_val(b) {}
    Value(double n) : type(Number), num_val(n) {}
    Value(const std::string& s) : type(String), str_val(s) {}
    Value(const char* s) : type(String), str_val(s) {}
    
    bool is_null() const { return type == Null; }
    bool is_bool() const { return type == Bool; }
    bool is_number() const { return type == Number; }
    bool is_string() const { return type == String; }
    bool is_array() const { return type == Array; }
    bool is_object() const { return type == Object; }
    
    bool as_bool() const { return type == Bool ? bool_val : false; }
    double as_double() const { return type == Number ? num_val : 0; }
    int as_int() const { return static_cast<int>(as_double()); }
    std::string as_string() const { 
        if (type == String) return str_val;
        if (type == Number) return std::to_string(num_val);
        if (type == Bool) return bool_val ? "true" : "false";
        return "";
    }
    
    Value& operator[](const std::string& key) {
        type = Object;
        return obj_val[key];
    }
    
    Value& operator[](size_t idx) {
        type = Array;
        if (idx >= arr_val.size()) arr_val.resize(idx + 1);
        return arr_val[idx];
    }
    
    const Value& operator[](const std::string& key) const {
        static Value null_val;
        auto it = obj_val.find(key);
        return it != obj_val.end() ? it->second : null_val;
    }
    
    bool has(const std::string& key) const {
        return type == Object && obj_val.find(key) != obj_val.end();
    }
    
    size_t size() const {
        return type == Array ? arr_val.size() : (type == Object ? obj_val.size() : 0);
    }
    
    std::vector<std::string> keys() const {
        std::vector<std::string> result;
        if (type == Object) {
            for (const auto& p : obj_val) result.push_back(p.first);
        }
        return result;
    }
};

class Parser {
    const char* p = nullptr;
    const char* end = nullptr;
    
    void skip_ws() { while (p < end && std::isspace(*p)) ++p; }
    
    char peek() { return p < end ? *p : 0; }
    char get() { return p < end ? *p++ : 0; }
    
    bool match(const char* s) {
        size_t len = strlen(s);
        if (p + len <= end && memcmp(p, s, len) == 0) { p += len; return true; }
        return false;
    }
    
    std::string parse_string() {
        std::string result;
        if (get() != '"') return result;
        while (p < end) {
            char c = get();
            if (c == '"') break;
            if (c == '\\' && p < end) {
                c = get();
                switch (c) {
                    case 'n': c = '\n'; break;
                    case 'r': c = '\r'; break;
                    case 't': c = '\t'; break;
                    case 'u': {
                        if (p + 4 <= end) {
                            char hex[5] = {p[0], p[1], p[2], p[3], 0};
                            p += 4;
                            int code = strtol(hex, nullptr, 16);
                            if (code < 0x80) {
                                result += (char)code;
                            } else if (code < 0x800) {
                                result += (char)(0xC0 | (code >> 6));
                                result += (char)(0x80 | (code & 0x3F));
                            } else {
                                result += (char)(0xE0 | (code >> 12));
                                result += (char)(0x80 | ((code >> 6) & 0x3F));
                                result += (char)(0x80 | (code & 0x3F));
                            }
                        }
                        continue;
                    }
                }
            }
            result += c;
        }
        return result;
    }
    
    Value parse_number() {
        const char* start = p;
        if (peek() == '-') ++p;
        while (p < end && std::isdigit(*p)) ++p;
        if (peek() == '.') { ++p; while (p < end && std::isdigit(*p)) ++p; }
        if (peek() == 'e' || peek() == 'E') {
            ++p;
            if (peek() == '+' || peek() == '-') ++p;
            while (p < end && std::isdigit(*p)) ++p;
        }
        return Value(std::stod(std::string(start, p)));
    }
    
public:
    Value parse(const char* str, size_t len) {
        p = str;
        end = str + len;
        return parse_value();
    }
    
    Value parse(const std::string& s) { return parse(s.c_str(), s.size()); }
    
    Value parse_value() {
        skip_ws();
        if (p >= end) return Value();
        
        char c = peek();
        if (c == 'n' && match("null")) return Value();
        if (c == 't' && match("true")) return Value(true);
        if (c == 'f' && match("false")) return Value(false);
        if (c == '"') return Value(parse_string());
        if (c == '-' || std::isdigit(c)) return parse_number();
        if (c == '[') return parse_array();
        if (c == '{') return parse_object();
        return Value();
    }
    
    Value parse_array() {
        Value arr;
        arr.type = Value::Array;
        get();
        skip_ws();
        if (peek() == ']') { get(); return arr; }
        while (true) {
            arr.arr_val.push_back(parse_value());
            skip_ws();
            if (peek() == ',') { get(); skip_ws(); }
            else if (peek() == ']') { get(); break; }
            else break;
        }
        return arr;
    }
    
    Value parse_object() {
        Value obj;
        obj.type = Value::Object;
        get();
        skip_ws();
        if (peek() == '}') { get(); return obj; }
        while (true) {
            skip_ws();
            if (peek() != '"') break;
            std::string key = parse_string();
            skip_ws();
            if (get() != ':') break;
            obj.obj_val[key] = parse_value();
            skip_ws();
            if (peek() == ',') { get(); }
            else if (peek() == '}') { get(); break; }
            else break;
        }
        return obj;
    }
};

inline Value parse(const std::string& s) {
    Parser p;
    return p.parse(s);
}

inline std::string stringify(const Value& v, int indent = 0) {
    std::ostringstream ss;
    std::string pad(indent, ' ');
    
    switch (v.type) {
        case Value::Null: ss << "null"; break;
        case Value::Bool: ss << (v.bool_val ? "true" : "false"); break;
        case Value::Number: ss << v.num_val; break;
        case Value::String: {
            ss << '"';
            for (char c : v.str_val) {
                switch (c) {
                    case '"': ss << "\\\""; break;
                    case '\\': ss << "\\\\"; break;
                    case '\n': ss << "\\n"; break;
                    case '\r': ss << "\\r"; break;
                    case '\t': ss << "\\t"; break;
                    default: ss << c;
                }
            }
            ss << '"';
            break;
        }
        case Value::Array:
            ss << "[";
            for (size_t i = 0; i < v.arr_val.size(); ++i) {
                if (i) ss << ", ";
                ss << stringify(v.arr_val[i], indent);
            }
            ss << "]";
            break;
        case Value::Object:
            ss << "{";
            if (!v.obj_val.empty()) ss << "\n";
            for (auto it = v.obj_val.begin(); it != v.obj_val.end(); ++it) {
                if (it != v.obj_val.begin()) ss << ",\n";
                ss << pad << "  \"" << it->first << "\": " << stringify(it->second, indent + 2);
            }
            if (!v.obj_val.empty()) ss << "\n" << pad;
            ss << "}";
            break;
    }
    return ss.str();
}

}
