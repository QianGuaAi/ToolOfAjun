#pragma once

#include <condition_variable>
#include <functional>
#include <mutex>
#include <queue>
#include <thread>

namespace mytools {

class TaskRunner {
public:
    TaskRunner();
    ~TaskRunner();

    TaskRunner(const TaskRunner&) = delete;
    TaskRunner& operator=(const TaskRunner&) = delete;

    void Post(std::function<void()> task);
    void Stop();

private:
    void WorkerLoop();

    std::mutex mutex_;
    std::condition_variable available_;
    std::queue<std::function<void()>> tasks_;
    std::thread worker_;
    bool stopping_ = false;
};

}  // namespace mytools
