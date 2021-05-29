var path = require("path");
var Q = require("q");
var fs = require("fs");

function copyTree(srcdir, destdir, callback) {
    fs.mkdir(destdir, function(err) {
        if (err && err.code != "EEXIST") throw new Error(err);
        fs.readdir(srcdir, function(err, files) {
            if (err) throw new Error(err);
            var count = files.length;

            function next(err) {
                if (err) throw new Error(err);
                if (--count == 0) callback();
            }
            if (count == 0) {
                callback();
            } else {
                files.forEach(function(f) {
                    var fullname = path.join(srcdir, f);
                    var dest = path.join(destdir, f);
                    fs.lstat(fullname, function(err, stat) {
                        if (err) throw new Error(err);
                        if (stat.isSymbolicLink()) {
                            fs.readlink(fullname, function(err, target) {
                                if (err) throw new Error(err);
                                fs.symlink(target, dest, next);
                            });
                        } else if (stat.isDirectory()) {
                            copyTree(fullname, dest, next);
                        } else if (stat.isFile()) {
                            fs.readFile(fullname, function(err, data) {
                                if (err) throw new Error(err);
                                fs.writeFile(dest, data, next);
                            });
                        } else {
                            next();
                        }
                    });
                });
            }
        });
    });
}

var productTasksDir = path.join(__dirname, "../ProductTasks");
var sharedFilesDir = path.join(__dirname, "../SharedFiles");

function getProductTaskDirs() {
    var productDirs = fs.readdirSync(productTasksDir).filter(function(file) {
        return fs.statSync(path.join(productTasksDir, file)).isDirectory();
    });

    return productDirs.map(function(productDir) {
        var fullProductDir = path.join(productTasksDir, productDir);

        return fs.readdirSync(fullProductDir).map(function(file) {
            return path.join(fullProductDir, file);
        }).filter(function(filePath) {
            return fs.statSync(filePath).isDirectory();
        });
    }).reduce((acc, val) => acc.concat(val), []);
}

var productTaskDirs = getProductTaskDirs();
var promisses = productTaskDirs.map(function(productTaskDir) {
    var deferred = Q.defer();

    console.log("Copying shared files to " + productTaskDir);

    copyTree(sharedFilesDir, productTaskDir, function() {
        console.log("Copy shared files done");
    });

    return deferred.promise;
});

Q.all(promisses)
    .fail(function(err) {
        console.error(err);
        process.exit(1);
    });