#include "services/task_runner.h"

namespace mytools {

TaskRunner::TaskRunner() : worker_([this]() { WorkerLoop(); }) {}

TaskRunner::~TaskRunner() {
    Stop();
}

void TaskRunner::Post(std::function<void()> task) {
    if (!task) {
        return;
    }

    {
        std::lock_guard<std::mutex> lock(mutex_);
        if (stopping_) {
            return;
        }
        tasks_.push(std::move(task));
    }
    available_.notify_one();
}

void TaskRunner::Stop() {
    {
        std::lock_guard<std::mutex> lock(mutex_);
        if (stopping_) {
            return;
        }
        stopping_ = true;
    }
    available_.notify_all();
    if (worker_.joinable()) {
        worker_.join();
    }
}

void TaskRunner::WorkerLoop() {
    for (;;) {
        std::function<void()> task;
        {
            std::unique_lock<std::mutex> lock(mutex_);
            available_.wait(lock, [this]() { return stopping_ || !tasks_.empty(); });
            if (stopping_ && tasks_.empty()) {
                return;
            }
            task = std::move(tasks_.front());
            tasks_.pop();
        }
        task();
    }
}

}  // namespace mytools
